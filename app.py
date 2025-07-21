import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QSystemTrayIcon, QMenu, QAction
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton, QLabel, QCheckBox, QStyle
from PyQt5.QtCore import QTimer, QThread, pyqtSignal

from ui_main_window import Ui_main_window # Corrected import based on your code

from win10toast import ToastNotifier
import schedule
import time
from datetime import datetime, timedelta

# --- Settings and Constants ---
# Path to the Excel file
def get_excel_path(filename):
    if hasattr(sys, '_MEIPASS'):
        # If packaged with PyInstaller, read the file from the temporary folder
        return os.path.join(sys._MEIPASS, filename)
    # Otherwise (when running from source code)
    return filename

EXCEL_FILE = get_excel_path('boss_takvimi.xlsx') # ## REMAINS TURKISH if EXCEL file column names are Turkish
BOSS_COLUMN = 'Boss' # Excel's boss name column (MUST EXACTLY MATCH your Excel header!)
DAY_COLUMN = 'Gun'   # Excel's day column (e.g., Pazartesi, Salı -> Monday, Tuesday). This is the header in Excel. ## REMAINS TURKISH
TIME_COLUMN = 'Saat' # Excel's time column (e.g., 01:00, 14:30). This is the header in Excel. ## REMAINS TURKISH

NOTIFICATION_TIMES = { # Notification times in minutes
    '1_min': 1,
    '3_min': 3,
    '5_min': 5,
    '10_min': 10,
    '15_min': 15,
    '30_min': 30,
    '60_min': 60
}

WEEKDAY_MAP = { # Map day names to numbers (Python's weekday() method takes Monday as 0, Sunday as 6)
    'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
    'Friday': 4, 'Saturday': 5, 'Sunday': 6,
    'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3, # Lowercase support
    'friday': 4, 'saturday': 5, 'sunday': 6
}

# --- Application Class ---
class BossNotificationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_main_window()
        self.ui.setupUi(self) # Setting up the UI on the main window

        # Load Style Sheet (if any)
        style_sheet_path = get_excel_path('style.qss')
        try:
            with open(style_sheet_path, "r") as f:
                self.setStyleSheet(f.read())
            print("style.qss loaded successfully.") ## TRANSLATED
        except FileNotFoundError:
            print(f"Warning: style.qss file not found: {style_sheet_path}") ## TRANSLATED
        except Exception as e:
            print(f"Warning: Error loading style.qss: {e}") ## TRANSLATED

        self.toast = ToastNotifier()
        self.active_bosses = {} # Stores active bosses selected by user {boss_name: True/False}
        self.notification_settings = { # Stores selected notification times {time_key: True/False}
            '1_min': False, '3_min': False, '5_min': False, '10_min': False,
            '15_min': False, '30_min': False, '60_min': False
        }
        self.df_schedule = None # Schedule data read from Excel
        self.notifications_sent_today = {} # Stores notifications sent today {event_id: [sent_times]}
        self._last_reset_date = None # Stores the last date notification history was reset

        self.initialize_ui_elements() # Connect UI elements
        self.load_boss_data() # Load boss data from Excel
        self.load_settings() # Load user settings (active bosses, notification times)
        self.apply_settings_to_ui() # Apply loaded settings to the UI

        # Create Scheduler Thread
        self.scheduler_thread = SchedulerThread(self)
        self.is_running = False # Indicates if the application's notification loop is running

        # System Tray Icon Settings
        self.tray_icon = QSystemTrayIcon(self)

        # Set application icon. If you don't have your own icon file, we can use a default system icon.
        # If you want to use your own icon file (e.g., 'app_icon.ico' in the same folder):
        # self.tray_icon.setIcon(QIcon(get_excel_path('app_icon.ico'))) ## CHANGE - use get_excel_path for icon too
        # Otherwise, let's use a standard system icon:
        icon = self.style().standardIcon(QStyle.SP_ComputerIcon) # Example system icon
        self.tray_icon.setIcon(icon)

        self.tray_icon.setToolTip("Boss Notification App") ## TRANSLATED

        # Create System Tray Menu
        tray_menu = QMenu()
        open_action = QAction("Open Application", self) ## TRANSLATED
        open_action.triggered.connect(self.showNormal) # Show the application again
        tray_menu.addAction(open_action)

        exit_action = QAction("Exit", self) ## TRANSLATED
        exit_action.triggered.connect(QApplication.instance().quit) # Completely close the application
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu) # Connect menu to the icon
        self.tray_icon.activated.connect(self.tray_icon_activated) # Connect icon click event

        self.update_status_label("Application ready. Select bosses to track and click Start.") ## TRANSLATED

    def initialize_ui_elements(self):
        """Connects UI elements to Python objects and assigns signals."""
        # Find and Connect Boss Buttons
        self.boss_buttons = {}
        all_boss_names = self.get_all_boss_names_from_excel_file() # Get all boss names from Excel

        for boss_name in all_boss_names:
            # Our objectName rule: boss_name_btn (e.g., garmoth_btn, golden_pig_king_btn)
            # If boss names contain spaces or special characters, normalize them for objectName
            button_name_normalized = boss_name.lower().replace(' ', '_').replace('-', '_').replace('.', '_')
            button_name = f"{button_name_normalized}_btn"

            button = self.findChild(QPushButton, button_name)
            if button:
                self.boss_buttons[boss_name] = button
                button.setCheckable(True) # Make buttons toggleable
                button.clicked.connect(lambda checked, b=boss_name: self.toggle_boss_active(b, checked))
                self.active_bosses[boss_name] = False # Initially all are marked inactive
            else:
                # Warn if a boss from Excel doesn't have a corresponding button in the UI
                print(f"Warning: A button named '{button_name}' was not found for boss '{boss_name}'. Please check your UI design.") ## TRANSLATED

        # Find and Connect Notification Time Checkboxes
        self.notification_checkboxes = {}
        for key, minutes in NOTIFICATION_TIMES.items():
            checkbox_name = f"chk_{minutes}min" # Our objectName rule: chk_1min, chk_3min etc.
            checkbox = self.findChild(QCheckBox, checkbox_name)
            if checkbox:
                self.notification_checkboxes[key] = checkbox
                checkbox.setChecked(False) # Initially all are inactive
                checkbox.stateChanged.connect(lambda state, k=key: self.toggle_notification_setting(k, state))
            else:
                print(f"Warning: A checkbox named '{checkbox_name}' was not found for '{minutes} Minutes Before' notification. Please check your UI design.") ## TRANSLATED

        # Connect Start/Stop Buttons
        self.ui.start_btn.clicked.connect(self.start_monitoring)
        self.ui.stop_btn.clicked.connect(self.stop_monitoring)
        self.ui.stop_btn.setEnabled(False) # Stop button initially disabled
        # Connect the newly added "Minimize to Tray" button
        self.ui.minimize_to_tray_btn.clicked.connect(self.minimize_to_tray)

        # Status Label
        self.status_label = self.ui.status_label # Should already be defined with this name in ui_main_window.py

    def get_all_boss_names_from_excel_file(self):
        """Extracts unique boss names from the Excel file."""
        try:
            temp_df = pd.read_excel(EXCEL_FILE)
            if BOSS_COLUMN not in temp_df.columns:
                QMessageBox.critical(self, "Error", f"'{BOSS_COLUMN}' column not found in Excel file.\nPlease check your Excel file. Header should be '{BOSS_COLUMN}'.") ## TRANSLATED
                sys.exit(1)
            return sorted(temp_df[BOSS_COLUMN].unique().tolist())
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", f"'{EXCEL_FILE}' file not found.\nPlease ensure it's in the same directory as the application.") ## TRANSLATED
            sys.exit(1)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while reading the Excel file: {e}\nPlease ensure the file is in the correct format (xlsx) and not open.") ## TRANSLATED
            sys.exit(1)

    def load_boss_data(self):
        """Loads the event schedule from the Excel file."""
        try:
            self.df_schedule = pd.read_excel(EXCEL_FILE)
            # Check for required columns
            required_columns = [DAY_COLUMN, TIME_COLUMN, BOSS_COLUMN]
            for col in required_columns:
                if col not in self.df_schedule.columns:
                    QMessageBox.critical(self, "Error", f"Column '{col}' not found in the Excel file.\nPlease check column headers (Gun, Saat, Boss).") ## TRANSLATED
                    sys.exit(1)

            # Convert Time column to datetime.time objects
            # We convert Excel's "HH:MM" text directly to time objects
            def parse_time_str(time_str):
                try:
                    # Try to convert the incoming value to string and then format it
                    return datetime.strptime(str(time_str).strip(), '%H:%M').time()
                except ValueError:
                    print(f"Warning: Invalid time format found: '{time_str}'. This row will be skipped.") ## TRANSLATED
                    return None # Skip rows with invalid format

            self.df_schedule['Parsed_Time'] = self.df_schedule[TIME_COLUMN].apply(parse_time_str)
            self.df_schedule = self.df_schedule.dropna(subset=['Parsed_Time']) # Remove rows with invalid times

            # Convert day names to lowercase to match WEEKDAY_MAP
            self.df_schedule['Normalized_Day'] = self.df_schedule[DAY_COLUMN].str.lower()

            self.update_status_label(f"'{os.path.basename(EXCEL_FILE)}' loaded successfully. {len(self.df_schedule)} valid events found.") ## TRANSLATED

        except FileNotFoundError:
            self.update_status_label(f"ERROR: '{os.path.basename(EXCEL_FILE)}' file not found. Please ensure it's in the correct directory.") ## TRANSLATED
            QMessageBox.critical(self, "Excel Error", f"'{os.path.basename(EXCEL_FILE)}' file not found.") ## TRANSLATED
            sys.exit(1)
        except Exception as e:
            self.update_status_label(f"ERROR: An issue occurred while loading the Excel file: {e}") ## TRANSLATED
            QMessageBox.critical(self, "Excel Error", f"An error occurred while loading the Excel file: {e}\nPlease ensure the file is in the correct format (xlsx) and not open.") ## TRANSLATED
            sys.exit(1)

    def save_settings(self):
        """Saves active bosses and notification settings (as a simple file)."""
        settings_path = 'settings.txt'
        try:
            with open(settings_path, 'w') as f:
                f.write("[Active Bosses]\n")
                for boss, active in self.active_bosses.items():
                    f.write(f"{boss}={active}\n")
                f.write("[Notification Settings]\n")
                for setting, active in self.notification_settings.items():
                    f.write(f"{setting}={active}\n")
            self.update_status_label("Settings saved.") ## TRANSLATED
        except Exception as e:
            self.update_status_label(f"Error saving settings: {e}") ## TRANSLATED

    def load_settings(self):
        """Loads saved settings."""
        settings_path = 'settings.txt'
        if not os.path.exists(settings_path):
            return # Do nothing if settings file doesn't exist

        try:
            with open(settings_path, 'r') as f:
                current_section = None
                for line in f:
                    line = line.strip()
                    if line.startswith('[') and line.endswith(']'):
                        current_section = line
                    elif current_section == '[Active Bosses]' and '=' in line:
                        parts = line.split('=', 1) # Split only by the first '='
                        if len(parts) == 2:
                            boss, active = parts
                            self.active_bosses[boss] = (active.lower() == 'true')
                    elif current_section == '[Notification Settings]' and '=' in line:
                        parts = line.split('=', 1) # Split only by the first '='
                        if len(parts) == 2:
                            setting, active = parts
                            self.notification_settings[setting] = (active.lower() == 'true')
        except Exception as e:
            self.update_status_label(f"An error occurred while loading settings: {e}") ## TRANSLATED

    def apply_settings_to_ui(self):
        """Applies loaded settings to the UI."""
        for boss, active in self.active_bosses.items():
            if boss in self.boss_buttons:
                self.boss_buttons[boss].setChecked(active)

        for key, active in self.notification_settings.items():
            if key in self.notification_checkboxes:
                self.notification_checkboxes[key].setChecked(active)

    def toggle_boss_active(self, boss_name, checked):
        """Changes the active status of a boss."""
        self.active_bosses[boss_name] = checked
        self.update_status_label(f"Tracking {boss_name}: {'Active' if checked else 'Inactive'}") ## TRANSLATED
        self.save_settings() # Save settings when changed

    def toggle_notification_setting(self, setting_key, state):
        """Changes the notification time setting."""
        self.notification_settings[setting_key] = bool(state)
        setting_text = self.notification_checkboxes[setting_key].text()
        self.update_status_label(f"{setting_text} notification: {'Active' if bool(state) else 'Inactive'}") ## TRANSLATED
        self.save_settings() # Save settings when changed

    def start_monitoring(self):
        """Starts the notification checking loop."""
        if self.is_running:
            self.update_status_label("Application is already running.") ## TRANSLATED
            return

        active_selected_bosses = [boss for boss, active in self.active_bosses.items() if active]
        if not active_selected_bosses:
            self.update_status_label("No bosses selected. Please select bosses you wish to track.") ## TRANSLATED
            QMessageBox.warning(self, "Warning", "Please select at least one boss.") ## TRANSLATED
            return

        active_notification_times = [
            NOTIFICATION_TIMES[key] for key, active in self.notification_settings.items() if active
        ]
        if not active_notification_times:
            self.update_status_label("No notification times selected. Please select a time.") ## TRANSLATED
            QMessageBox.warning(self, "Warning", "Please select at least one notification time (e.g., 15 Minutes Before).") ## TRANSLATED
            return

        self.is_running = True
        self.update_status_label("Notifications started.") ## TRANSLATED
        self.ui.start_btn.setEnabled(False) # Disable Start button
        self.ui.stop_btn.setEnabled(True) # Enable Stop button
        self.ui.minimize_to_tray_btn.setEnabled(True) ## CHANGE: Enable minimize to tray button when monitoring starts

        # Start/Resume Scheduler
        if not self.scheduler_thread.isRunning():
            self.scheduler_thread.start() # Start thread if not already running
        self.scheduler_thread.resume_scheduling() # Resume scheduling

    def stop_monitoring(self):
        """Stops the notification checking loop."""
        if not self.is_running:
            self.update_status_label("Application is already stopped.") ## TRANSLATED
            return

        self.is_running = False
        self.update_status_label("Notifications stopped.") ## TRANSLATED
        self.ui.start_btn.setEnabled(True) # Enable Start button
        self.ui.stop_btn.setEnabled(False) # Disable Stop button
        self.ui.minimize_to_tray_btn.setEnabled(False) ## CHANGE: Disable minimize to tray button when monitoring stops

        if self.scheduler_thread.isRunning():
            self.scheduler_thread.pause_scheduling() # Pause scheduling

    def minimize_to_tray(self):
        """Minimizes the application to the system tray."""
        self.hide() # Hide the main window
        self.tray_icon.show() # Show the system tray icon
        self.tray_icon.showMessage(
        "Boss Notification App", ## TRANSLATED
        "Application is running in the background. You will continue to receive notifications.", ## TRANSLATED
        self.tray_icon.icon(), # Show application icon in the message box
        2000 # Make the message visible for 2 seconds
        )
        self.update_status_label("Application minimized to tray.") ## TRANSLATED

    def tray_icon_activated(self, reason):
        """Manages the event when the system tray icon is clicked."""
        if reason == QSystemTrayIcon.Trigger or reason == QSystemTrayIcon.DoubleClick:
            self.showNormal() # Show the window in normal size
            self.activateWindow() # Bring the window to the front
            self.tray_icon.hide() # Hide the icon from the system tray
            self.update_status_label("Application window restored.") ## TRANSLATED

    def closeEvent(self, event):
        """Captures the window close button event."""
        event.ignore() # Prevent default close behavior
        self.minimize_to_tray() # Minimize the application to the system tray

    def check_for_notifications(self):
        """Checks for events and sends notifications."""
        if not self.is_running or self.df_schedule is None or self.df_schedule.empty:
            return

        now = datetime.now()
        current_weekday_num = now.weekday() # Monday=0, Sunday=6

        # Daily reset: Reset sent notifications when a new day begins after midnight
        today_date = now.date()
        if not hasattr(self, '_last_reset_date') or self._last_reset_date != today_date:
            self.notifications_sent_today = {}
            self._last_reset_date = today_date
            self.update_status_label(f"New day: {today_date.strftime('%Y-%m-%d')}. Notification history reset.") ## TRANSLATED

        # Filter only active boss events
        active_bosses_list = [boss for boss, active in self.active_bosses.items() if active]
        if not active_bosses_list:
            self.update_status_label("No active bosses found to track.") ## TRANSLATED
            return # Don't check if no active bosses

        # Filter using normalized day names from Excel
        possible_day_names = [k for k, v in WEEKDAY_MAP.items() if v == current_weekday_num]

        # Filter DataFrame by day and active bosses
        filtered_df = self.df_schedule[
            (self.df_schedule['Normalized_Day'].isin(possible_day_names)) &
            (self.df_schedule[BOSS_COLUMN].isin(active_bosses_list))
        ].copy()

        active_notification_times = [
            NOTIFICATION_TIMES[key] for key, active in self.notification_settings.items() if active
        ]

        next_event_info = None
        min_time_diff = timedelta(days=365) # Initialize with a very large value

        for index, row in filtered_df.iterrows():
            event_time_obj = row['Parsed_Time']
            event_datetime = datetime.combine(now.date(), event_time_obj)

            # If the event has passed (with a few seconds tolerance), move it to the next week
            if event_datetime < now - timedelta(seconds=10): # Consider it passed if 10 seconds ago
                event_datetime += timedelta(weeks=1)
                # If it's still in the past even after moving to next week (e.g., midnight boss in schedule),
                # this is a rare case, we'll leave it as is for now.

            time_diff = event_datetime - now

            # Find the next upcoming event
            # Only check future events and find the closest one
            if time_diff > timedelta(seconds=0) and time_diff < min_time_diff:
                next_event_info = (row[BOSS_COLUMN], event_datetime)
                min_time_diff = time_diff

            # Notification sending logic
            # Sort notification times from largest to smallest to catch the earliest one to send
            for notify_minutes in sorted(active_notification_times, reverse=True):
                notification_threshold = timedelta(minutes=notify_minutes)

                # Create a unique notification ID (using Day, Time, and Boss name)
                # Ensure a notification is sent only once for a specific event at a specific time
                event_id = f"{row[BOSS_COLUMN]}_{row[DAY_COLUMN]}_{row[TIME_COLUMN]}"
                notification_key = f"{event_id}_{notify_minutes}min"

                # Notification interval: e.g., for 15 minutes threshold, send notification if remaining time is between 15:00 and 14:00
                if notification_threshold >= time_diff > (notification_threshold - timedelta(seconds=30)) and \
                   notification_key not in self.notifications_sent_today:

                    title = f"Boss Notification: {row[BOSS_COLUMN]}" ## TRANSLATED
                    message = f"{row[BOSS_COLUMN]} will spawn in {notify_minutes} minutes, on {row[DAY_COLUMN]} at {row[TIME_COLUMN]}!" ## TRANSLATED
                    self.toast.show_toast(title, message, duration=10, icon_path=None, threaded=True)
                    self.notifications_sent_today[notification_key] = True # Mark this notification as sent
                    self.update_status_label(f"Notification Sent: {message}") ## TRANSLATED
                    break # Notification sent for this threshold for this event, no need to check other thresholds

        # Show next event info in the status bar
        if next_event_info:
            boss_name, next_time = next_event_info
            remaining_time = next_time - now
            hours, remainder = divmod(remaining_time.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)

            # If there's a difference in days, include day information
            if remaining_time.days > 0:
                self.update_status_label(f"Next event: {boss_name} - {next_time.strftime('%Y-%m-%d %H:%M')} ({remaining_time.days} days {hours}h {minutes}min)") ## TRANSLATED
            else:
                self.update_status_label(f"Next event: {boss_name} - {next_time.strftime('%H:%M')} ({hours}h {minutes}min)") ## TRANSLATED
        else:
            self.update_status_label("No upcoming events found.") ## TRANSLATED


    def update_status_label(self, message):
        """Updates the status label (updated on the GUI main thread)."""
        # Using pyqtSignal for safe inter-thread communication
        self.ui.status_label.setText(message)

# --- Scheduler Thread (Runs in Background) ---
class SchedulerThread(QThread):
    def __init__(self, app_instance):
        super().__init__()
        self.app = app_instance
        self._is_paused = True # Initially paused
        self._stop_event = False # To completely stop the thread
        # Check every 30 seconds. You can change this interval as needed.
        schedule.every(30).seconds.do(self.app.check_for_notifications)

    def run(self):
        while not self._stop_event:
            if not self._is_paused:
                schedule.run_pending()
            time.sleep(1) # Check schedule every second

    def pause_scheduling(self):
        """Pauses the scheduling loop."""
        self._is_paused = True

    def resume_scheduling(self):
        """Resumes the scheduling loop."""
        self._is_paused = False

    def stop_thread(self):
        """Completely stops the thread."""
        self._stop_event = True
        schedule.clear() # Clear all scheduled jobs
        self.wait() # Wait for the thread to finish (for safe shutdown)

    def __del__(self):
        # Ensures thread is cleaned up when application closes
        if self.isRunning():
            self.stop_thread()


# --- Run the Application ---
if __name__ == '__main__':
    # QApplication creates the base of the application.
    # sys.argv passes command-line arguments to PyQt.
    app = QApplication(sys.argv)

    # Create an instance of our main application class
    main_app = BossNotificationApp()

    # Show the main window
    main_app.show()

    # Starts the application's event loop.
    # sys.exit() ensures the application closes properly.
    sys.exit(app.exec_())