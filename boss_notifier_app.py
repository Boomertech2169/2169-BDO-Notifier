import sys
import json
import time
from datetime import datetime, timedelta
from plyer import notification
import os
from PIL import Image # Pillow kütüphanesi

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QScrollArea, QCheckBox, QPushButton, QLabel, QGroupBox, QSizePolicy, QAction, QMenu,
    QStyle # <-- Bu satırı ekle
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon # PyQt5'te ikonlar için QIcon kullanılır
from PyQt5.QtWidgets import QSystemTrayIcon # Sistem tepsisi için PyQt'nin kendi aracı

# --- Yeni Ekleyeceğin Bölüm Başlangıcı ---
def get_resource_path(relative_path):
    """
    Kaynak dosyalarının (JSON, ikon vb.) hem geliştirme ortamında hem de PyInstaller ile
    paketlendiğinde doğru yolunu döndürür.
    """
    try:
        # PyInstaller çalıştırılabilir dosya oluşturduğunda _MEIPASS özniteliğini ekler.
        base_path = sys._MEIPASS
    except Exception:
        # Geliştirme ortamında çalışıyorsa geçerli dizini kullan.
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# JSON dosyasının yolu
BOSS_DATA_FILE = get_resource_path('boss_data.json')
# İkon dosyası yolu
ICON_PATH = get_resource_path('icon.png')
# --- Yeni Ekleyeceğin Bölüm Sonu ---

# Gün isimlerini Python'ın datetime modülünün anlayacağı şekilde eşleştirme
DAY_MAPPING = {
    "Monday": 0,
    "Tuesday": 1,
    "Wednesday": 2,
    "Thursday": 3,
    "Friday": 4,
    "Saturday": 5,
    "Sunday": 6
}

# --- Yardımcı Fonksiyonlar ---
def load_boss_data(file_path):
    """JSON dosyasından boss verilerini yükler."""
    if not os.path.exists(file_path):
        print(f"Hata: '{file_path}' dosyası bulunamadı. Lütfen dosyanın uygulamanın çalıştığı dizinde olduğundan emin olun.")
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data
    except json.JSONDecodeError:
        print(f"Hata: '{file_path}' dosyası geçersiz bir JSON formatına sahip.")
        return None
    except Exception as e:
        print(f"Dosya okunurken bir hata oluştu: {e}")
        return None

def send_notification(title, message, app_name="Boss Notifier", timeout=10, toast=True):
    """Windows bildirimi gönderir (plyer kullanır)."""
    try:
        notification.notify(
            title=title,
            message=message,
            app_name=app_name,
            timeout=timeout,
            toast=toast
        )
    except Exception as e:
        print(f"Bildirim gönderilirken bir hata oluştu: {e}")
        print("Plyer kütüphanesi doğru yüklenmemiş veya sistemde bir sorun olabilir.")

def get_next_spawn_time(boss_name, spawn_times, current_time):
    """Bir boss'un bir sonraki spawn zamanını bulur."""
    next_spawn = None
    current_weekday = current_time.weekday()

    for spawn in spawn_times:
        day_of_week_int = DAY_MAPPING.get(spawn['day'])
        if day_of_week_int is None:
            continue

        hour, minute = map(int, spawn['time'].split(':'))
        
        days_diff = (day_of_week_int - current_weekday + 7) % 7

        spawn_datetime_candidate = current_time + timedelta(days=days_diff)
        spawn_datetime_candidate = spawn_datetime_candidate.replace(hour=hour, minute=minute, second=0, microsecond=0)

        if spawn_datetime_candidate <= current_time:
            spawn_datetime_candidate += timedelta(weeks=1)

        if next_spawn is None or spawn_datetime_candidate < next_spawn:
            next_spawn = spawn_datetime_candidate
            
    return next_spawn

# --- Arka Plan Bildirim Kontrol Thread'i ---
class NotificationWorker(QThread):
    # Ana pencereye güncel bilgiyi göndermek için sinyal
    live_info_updated = pyqtSignal(str)
    # Bildirim göndermek için sinyal (main thread'den notification.notify'i çağırırız)
    send_notification_signal = pyqtSignal(str, str)

    def __init__(self, boss_data, selected_bosses_callback, selected_minutes_callback):
        super().__init__()
        self.boss_data = boss_data
        self.get_selected_bosses = selected_bosses_callback
        self.get_selected_minutes = selected_minutes_callback
        self._is_active = False
        self.notified_spawn_times = set()

    def set_active(self, active):
        self._is_active = active

    def run(self):
        while True:
            current_time = datetime.now()
            
            # --- Canlı Bilgi Güncelleme Mantığı ---
            # Sadece uygulama aktifse veya güncel bilgi göstermek için her zaman kontrol et
            closest_bosses_info = []
            overall_earliest_spawn = None
            
            selected_boss_ids = self.get_selected_bosses() # Ana thread'den seçili boss'ları al
            
            for boss in self.boss_data:
                boss_id = boss['id']
                if boss_id not in selected_boss_ids:
                    continue
                
                if 'spawn_times' not in boss or not boss['spawn_times']:
                    continue

                next_spawn_for_boss = get_next_spawn_time(boss['name'], boss['spawn_times'], current_time)

                if next_spawn_for_boss:
                    if overall_earliest_spawn is None or next_spawn_for_boss < overall_earliest_spawn:
                        overall_earliest_spawn = next_spawn_for_boss
                        closest_bosses_info = [(boss['name'], next_spawn_for_boss)]
                    elif next_spawn_for_boss == overall_earliest_spawn:
                        closest_bosses_info.append((boss['name'], next_spawn_for_boss))
            
            live_info_text = ""
            if closest_bosses_info:
                if len(closest_bosses_info) > 1:
                    live_info_text += "En yakın bosslar:\n"
                    for name, spawn_time in closest_bosses_info:
                        time_until = spawn_time - current_time
                        total_seconds = int(time_until.total_seconds())
                        if total_seconds <= 0:
                            remaining_str = "ŞİMDİ!"
                        else:
                            hours, remainder = divmod(total_seconds, 3600)
                            minutes, seconds = divmod(remainder, 60)
                            remaining_str = f"{hours:02d}sa {minutes:02d}dk"
                        live_info_text += f" - {name} ({remaining_str})\n"
                else:
                    name, spawn_time = closest_bosses_info[0]
                    time_until = spawn_time - current_time
                    total_seconds = int(time_until.total_seconds())
                    if total_seconds <= 0:
                        remaining_str = "ŞİMDİ!"
                    else:
                        hours, remainder = divmod(total_seconds, 3600)
                        minutes, seconds = divmod(remainder, 60)
                        remaining_str = f"{hours:02d}sa {minutes:02d}dk"
                    live_info_text = f"En yakın boss: {name} ({remaining_str})"
            else:
                live_info_text = "Takip edilecek boss bulunamadı veya tüm bosslar seçili değil."
            
            self.live_info_updated.emit(live_info_text) # Ana thread'e sinyal gönder

            # --- Bildirim Gönderme Mantığı ---
            if self._is_active: # Sadece uygulama aktifse bildirimleri kontrol et
                selected_minutes = self.get_selected_minutes() # Ana thread'den seçili dakikaları al
                for boss in self.boss_data:
                    boss_id = boss['id']
                    boss_name = boss['name']

                    if boss_id not in selected_boss_ids: # Tekrar kontrol
                        continue
                    
                    if 'spawn_times' not in boss or not boss['spawn_times']:
                        continue

                    next_spawn = get_next_spawn_time(boss_name, boss['spawn_times'], current_time)

                    if next_spawn:
                        time_until_spawn = next_spawn - current_time

                        for minutes_before in selected_minutes:
                            notification_threshold = timedelta(minutes=minutes_before)
                            
                            if notification_threshold >= time_until_spawn > (notification_threshold - timedelta(minutes=1, seconds=5)):
                                
                                notification_key = (boss_id, next_spawn, minutes_before)
                                if notification_key not in self.notified_spawn_times:
                                    
                                    remaining_minutes = int(time_until_spawn.total_seconds() / 60)
                                    
                                    if remaining_minutes > 0:
                                        message_suffix = f"{remaining_minutes} dakika içinde çıkacak!"
                                    else:
                                        message_suffix = "şimdi çıkıyor!"
                                    
                                    notification_message = f"{boss_name} {message_suffix} (Tahmini: {next_spawn.strftime('%H:%M')})"
                                    
                                    self.send_notification_signal.emit(f"Boss Yaklaşıyor: {boss_name}", notification_message)
                                    print(f"Bildirim gönderildi: {notification_message}")
                                    self.notified_spawn_times.add(notification_key)
            
            # Geçmiş spawn zamanlarını notified_spawn_times setinden temizle
            self.notified_spawn_times = {
                (b_id, s_time, m_before) for b_id, s_time, m_before in self.notified_spawn_times
                if s_time > (current_time - timedelta(minutes=15))
            }

            time.sleep(15) # Her 10 saniyede bir kontrol et

# --- Ana Uygulama Penceresi ---
class BossNotifierApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Boss Notifier")
        self.setGeometry(100, 100, 650, 550) # x, y, width, height
        self.setMinimumSize(600, 500)

        self.boss_data = load_boss_data(BOSS_DATA_FILE)
        if self.boss_data is None:
            QApplication.quit()
            return

        self.selected_bosses_checkboxes = {} # {boss_id: QCheckBox}
        self.selected_notification_minutes_checkboxes = {} # {minute: QCheckBox}

        self.create_widgets()
        self.setup_system_tray()

        # NotificationWorker thread'ini başlat
        self.worker_thread = NotificationWorker(
            self.boss_data, 
            self.get_selected_boss_ids, 
            self.get_selected_notification_minutes
        )
        self.worker_thread.live_info_updated.connect(self.update_live_info_label)
        self.worker_thread.send_notification_signal.connect(send_notification) # Bildirimi main thread'den gönder
        self.worker_thread.start() # Thread'i başlat

        # Uygulama kapatıldığında (X tuşu) pencereyi gizle
        self.closeEvent = self.hide_window_to_tray

    def create_widgets(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget) # Ana dikey düzen

        # Üst kısım: Boss ve Dakika Seçimleri (yatay düzen)
        top_layout = QHBoxLayout()
        
        # --- Sol Bölüm: Boss Seçimi ---
        boss_group_box = QGroupBox("Takip Edilecek Bosslar:")
        boss_group_box.setStyleSheet("QGroupBox { font-weight: bold; }")
        boss_layout = QVBoxLayout(boss_group_box)
        
        # Kaydırılabilir alan için QScrollArea
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content_widget = QWidget()
        self_scroll_content_layout = QVBoxLayout(scroll_content_widget)
        self_scroll_content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft) # Checkbox'ları üste ve sola hizala
        scroll_area.setWidget(scroll_content_widget)

        for boss in self.boss_data:
            boss_id = boss['id']
            cb = QCheckBox(boss['name'])
            cb.setChecked(True) # Varsayılan olarak hepsi seçili
            self.selected_bosses_checkboxes[boss_id] = cb
            self_scroll_content_layout.addWidget(cb)
        
        # Boss layout'a scroll area ekle
        boss_layout.addWidget(scroll_area)
        top_layout.addWidget(boss_group_box, 3) # Boss bölümü daha geniş olsun (weight 3)

        # --- Sağ Bölüm: Bildirim Dakikası Seçimi ---
        minutes_group_box = QGroupBox("Bildirim Süresi (Dakika Önce):")
        minutes_group_box.setStyleSheet("QGroupBox { font-weight: bold; }")
        minutes_layout = QVBoxLayout(minutes_group_box)
        minutes_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft) # Checkbox'ları üste ve sola hizala

        notification_minutes_options = [1, 3, 5, 10, 15, 30]
        for minute in notification_minutes_options:
            cb = QCheckBox(f"{minute} dakika")
            cb.setChecked(False) # Varsayılan olarak hiçbiri seçili gelmesin
            self.selected_notification_minutes_checkboxes[minute] = cb
            minutes_layout.addWidget(cb)
        
        top_layout.addWidget(minutes_group_box, 1) # Dakika bölümü normal genişlikte (weight 1)

        main_layout.addLayout(top_layout)

        # --- Canlı Bilgi Alanı (Ortaya Eklendi) ---
        self.live_info_label = QLabel("Bilgiler yükleniyor...")
        self.live_info_label.setAlignment(Qt.AlignCenter) # Metni ortala
        self.live_info_label.setStyleSheet("""
            QLabel {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
        """)
        self.live_info_label.setWordWrap(True) # Metnin sığmaması durumunda alt satıra geç
        self.live_info_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # Yatayda genişlesin, dikeyde sabit
        main_layout.addWidget(self.live_info_label)

        # --- Kontrol Butonları Bölümü (Alt Kısım) ---
        control_layout = QHBoxLayout()
        self.active_button = QPushButton("Active")
        self.active_button.clicked.connect(self.activate_notifications)
        self.active_button.setEnabled(True) # Varsayılan aktif

        self.deactive_button = QPushButton("Deactive")
        self.deactive_button.clicked.connect(self.deactivate_notifications)
        self.deactive_button.setEnabled(False) # Varsayılan deaktif

        control_layout.addWidget(self.active_button)
        control_layout.addWidget(self.deactive_button)
        main_layout.addLayout(control_layout)

    def get_selected_boss_ids(self):
        """Seçili boss ID'lerini döndürür."""
        return [boss_id for boss_id, cb in self.selected_bosses_checkboxes.items() if cb.isChecked()]

    def get_selected_notification_minutes(self):
        """Seçili bildirim dakikalarını döndürür."""
        return [minute for minute, cb in self.selected_notification_minutes_checkboxes.items() if cb.isChecked()]

    def update_live_info_label(self, text):
        """Canlı bilgi etiketini günceller (worker thread'den sinyal ile çağrılır)."""
        self.live_info_label.setText(text)

    def activate_notifications(self):
        """Bildirim gönderme işlemini başlatır ve buton durumlarını günceller."""
        self.worker_thread.set_active(True)
        self.active_button.setEnabled(False)
        self.deactive_button.setEnabled(True)
        print("Bildirimler Aktif.")

    def deactivate_notifications(self):
        """Bildirim gönderme işlemini durdurur ve buton durumlarını günceller."""
        self.worker_thread.set_active(False)
        self.active_button.setEnabled(True)
        self.deactive_button.setEnabled(False)
        print("Bildirimler Deaktif.")
        self.live_info_label.setText("Bildirimler deaktif. Bilgi güncellenmiyor.")

    def setup_system_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        
        # *** BURAYI DEĞİŞTİRİYORUZ ***
        # Standart bir sistem ikonu kullanmayı dene. Bu ikon her Windows sisteminde bulunur.
        self.tray_icon.setIcon(QApplication.style().standardIcon(QStyle.SP_ComputerIcon)) 
        # Ya da başka bir standart ikon deneyebilirsin: QStyle.SP_MessageBoxInformation, QStyle.SP_MediaPlay
        
        self.tray_icon.setToolTip("Boss Notifier")

        # Sistem tepsisi menüsü (bu kısım değişmedi)
        tray_menu = QMenu()
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show_window_from_tray)
        tray_menu.addAction(show_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show() # İkonu göster
        print("Sistem tepsisi ikonu gösterildi (standart ikon ile).") # Ek doğrulama

    def hide_window_to_tray(self, event=None):
        """Pencereyi gizler ve sistem tepsisine gönderir."""
        self.hide() # Pencereyi gizle
        print("Uygulama sistem tepsisine gizlendi.")
        # event'i kabul et ama kullanma, QCloseEvent'ten gelebilir

    def show_window_from_tray(self):
        """Gizlenmiş pencereyi tekrar gösterir."""
        self.showNormal() # Pencereyi normal boyutunda göster
        self.activateWindow() # Pencereyi öne getir
        print("Uygulama penceresi gösterildi.")

    def quit_app(self):
        """Uygulamayı tamamen kapatır."""
        self.tray_icon.hide() # Sistem tepsisi ikonunu gizle
        self.worker_thread.quit() # Thread'i güvenli bir şekilde sonlandır
        self.worker_thread.wait() # Thread'in bitmesini bekle
        QApplication.quit() # PyQt uygulamasını kapat


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Bu satırı buraya ekle:
    app.setQuitOnLastWindowClosed(False) 
    
    # Uygulama ikonunu ayarla (ana pencere ve görev çubuğu için)
    if os.path.exists(ICON_PATH):
        app.setWindowIcon(QIcon(ICON_PATH))

    main_window = BossNotifierApp()
    main_window.show() # Uygulama başlangıçta görünür olsun

    sys.exit(app.exec_())