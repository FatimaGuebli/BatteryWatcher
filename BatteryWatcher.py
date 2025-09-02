import sys
import threading
import time
import os
import psutil
from win10toast import ToastNotifier
import win32api
import win32con
import win32event
import win32gui
import pythoncom
from PyQt5.QtWidgets import (QApplication, QMainWindow, QSystemTrayIcon, 
                             QMenu, QAction, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QSpinBox, QCheckBox, QPushButton, QMessageBox, QStyle)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

# Single instance management
def is_app_already_running():
    """Check if another instance is already running"""
    try:
        # Create a mutex to check for existing instances
        mutex = win32event.CreateMutex(None, False, "BatteryWatcherMutex")
        return win32api.GetLastError() == win32con.ERROR_ALREADY_EXISTS
    except:
        return False

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class BatteryWatcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.tray_icon = None
        self.monitor_thread = None
        self.stop_monitoring = False
        self.toaster = ToastNotifier()
        self.min_level = 55
        self.max_level = 98
        self.auto_start = False

        # Load settings
        self.load_settings()

        # Setup UI
        self.init_ui()

        # Setup system tray
        self.setup_tray()

        # Start monitoring
        self.start_monitoring()

        # Set window properties
        self.setWindowTitle("BatteryWatcher")
        self.setFixedSize(300, 250)

        try:
            self.setWindowIcon(QIcon(resource_path("icon.ico")))
        except:
            pass

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        min_layout = QHBoxLayout()
        min_layout.addWidget(QLabel("Minimum Battery Level:"))
        self.min_spin = QSpinBox()
        self.min_spin.setRange(1, 100)
        self.min_spin.setValue(self.min_level)
        self.min_spin.valueChanged.connect(self.update_min_level)
        min_layout.addWidget(self.min_spin)
        layout.addLayout(min_layout)

        max_layout = QHBoxLayout()
        max_layout.addWidget(QLabel("Maximum Battery Level:"))
        self.max_spin = QSpinBox()
        self.max_spin.setRange(1, 100)
        self.max_spin.setValue(self.max_level)
        self.max_spin.valueChanged.connect(self.update_max_level)
        max_layout.addWidget(self.max_spin)
        layout.addLayout(max_layout)

        self.auto_start_cb = QCheckBox("Start automatically when PC turns on")
        self.auto_start_cb.setChecked(self.auto_start)
        self.auto_start_cb.stateChanged.connect(self.toggle_auto_start)
        layout.addWidget(self.auto_start_cb)

        close_button = QPushButton("Close to System Tray")
        close_button.clicked.connect(self.hide)
        layout.addWidget(close_button)

        exit_button = QPushButton("Exit Application")
        exit_button.clicked.connect(self.quit_app)
        layout.addWidget(exit_button)

    def setup_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            QMessageBox.critical(None, "System Tray", "System tray is not available on this system.")
            return

        try:
            tray_icon = QIcon(resource_path("icon.ico"))
        except:
            tray_icon = self.style().standardIcon(QStyle.SP_ComputerIcon)

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(tray_icon)
        self.tray_icon.setToolTip("BatteryWatcher - Monitoring your battery")

        tray_menu = QMenu()

        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self.show)
        tray_menu.addAction(settings_action)

        tray_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "BatteryWatcher",
            "Application minimized to system tray. It will continue monitoring your battery.",
            QSystemTrayIcon.Information,
            2000
        )

    def update_min_level(self, value):
        self.min_level = value
        self.save_settings()

    def update_max_level(self, value):
        self.max_level = value
        self.save_settings()

    def toggle_auto_start(self, state):
        self.auto_start = state == Qt.Checked
        self.set_auto_start(self.auto_start)
        self.save_settings()

    def load_settings(self):
        """Load settings from registry, fallback to defaults"""
        try:
            key = win32con.HKEY_CURRENT_USER
            key_path = r"Software\BatteryWatcher"
            reg_key = win32api.RegCreateKeyEx(
                key, key_path, 0, win32con.REG_OPTION_NON_VOLATILE, 
                win32con.KEY_READ | win32con.KEY_WRITE, None, None
            )

            try:
                min_val, _ = win32api.RegQueryValueEx(reg_key, "MinLevel")
                max_val, _ = win32api.RegQueryValueEx(reg_key, "MaxLevel")
                auto_val, _ = win32api.RegQueryValueEx(reg_key, "AutoStart")

                self.min_level = int(min_val)
                self.max_level = int(max_val)
                self.auto_start = bool(auto_val)
            except FileNotFoundError:
                self.min_level = 96
                self.max_level = 100
                self.auto_start = False

            win32api.RegCloseKey(reg_key)
            print(f"Loaded settings: Min={self.min_level}, Max={self.max_level}, AutoStart={self.auto_start}")
        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self):
        try:
            key = win32con.HKEY_CURRENT_USER
            key_path = r"Software\BatteryWatcher"
            reg_key = win32api.RegCreateKeyEx(
                key, key_path, 0, win32con.REG_OPTION_NON_VOLATILE,
                win32con.KEY_WRITE | win32con.KEY_READ, None, None
            )

            win32api.RegSetValueEx(reg_key, "MinLevel", 0, win32con.REG_DWORD, self.min_level)
            win32api.RegSetValueEx(reg_key, "MaxLevel", 0, win32con.REG_DWORD, self.max_level)
            win32api.RegSetValueEx(reg_key, "AutoStart", 0, win32con.REG_DWORD, int(self.auto_start))

            win32api.RegCloseKey(reg_key)
            print(f"Saved settings: Min={self.min_level}, Max={self.max_level}, AutoStart={self.auto_start}")
        except Exception as e:
            print(f"Error saving settings: {e}")

    def set_auto_start(self, enable):
        try:
            key = win32con.HKEY_CURRENT_USER
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

            reg_key = win32api.RegOpenKeyEx(key, key_path, 0, win32con.KEY_WRITE)
            if enable:
                exe_path = sys.executable if not getattr(sys, 'frozen', False) else sys.argv[0]
                win32api.RegSetValueEx(reg_key, "BatteryWatcher", 0, win32con.REG_SZ, f'"{exe_path}"')
                print("Added to startup")
            else:
                try:
                    win32api.RegDeleteValue(reg_key, "BatteryWatcher")
                    print("Removed from startup")
                except FileNotFoundError:
                    pass
            win32api.RegCloseKey(reg_key)
        except Exception as e:
            print(f"Error setting auto start: {e}")

    def start_monitoring(self):
        self.stop_monitoring = False
        self.monitor_thread = threading.Thread(target=self.monitor_battery)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()

    def stop_monitor(self):
        self.stop_monitoring = True
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=1)

    def monitor_battery(self):
        last_plugged = None
        last_notified = {"plugged": 0, "unplugged": 0}

        while not self.stop_monitoring:
            try:
                battery = psutil.sensors_battery()
                if battery is None:
                    print("No battery detected")
                    time.sleep(30)
                    continue

                plugged = battery.power_plugged
                percent = battery.percent

                print(f"Battery: {percent}%, Plugged: {plugged}, Min: {self.min_level}, Max: {self.max_level}")

                if plugged != last_plugged:
                    last_plugged = plugged
                    last_notified = {"plugged": 0, "unplugged": 0}
                    print(f"Power status changed: {'Plugged' if plugged else 'Unplugged'}")

                current_time = time.time()

                if plugged and percent >= self.max_level:
                    if current_time - last_notified["plugged"] > 300:
                        print(f"Notification: Battery fully charged ({percent}%)")
                        self.show_notification(
                            "Battery Fully Charged",
                            f"Battery is at {percent}%. Please unplug your charger to preserve battery health."
                        )
                        last_notified["plugged"] = current_time
                elif not plugged and percent <= self.min_level:
                    if current_time - last_notified["unplugged"] > 300:
                        print(f"Notification: Battery low ({percent}%)")
                        self.show_notification(
                            "Battery Low",
                            f"Battery is at {percent}%. Please plug in your charger."
                        )
                        last_notified["unplugged"] = current_time

            except Exception as e:
                print(f"Error monitoring battery: {e}")

            for _ in range(30):
                if self.stop_monitoring:
                    break
                time.sleep(1)

    def show_notification(self, title, message):
        try:
            self.toaster.show_toast(title, message, duration=10, threaded=False)
            print(f"Showing toast notification: {title}")
        except Exception as e:
            print(f"Error showing toast notification: {e}")
            if self.tray_icon:
                try:
                    self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 5000)
                    print("Fallback tray notification used")
                except Exception as e2:
                    print(f"Error showing tray notification: {e2}")

    def quit_app(self):
        reply = QMessageBox.question(self, 'Exit Confirmation', 
                                     "Are you sure you want to exit BatteryWatcher?",
                                     QMessageBox.Yes | QMessageBox.No, 
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.stop_monitor()
            if self.tray_icon:
                self.tray_icon.hide()
            QApplication.quit()

def handle_shutdown(signal, frame):
    print("System shutdown detected, exiting gracefully...")
    sys.exit(0)

def main():
    if is_app_already_running():
        try:
            def enum_windows_callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd) and "BatteryWatcher" in win32gui.GetWindowText(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
            win32gui.EnumWindows(enum_windows_callback, None)
        except:
            pass
        QMessageBox.warning(None, "BatteryWatcher", "BatteryWatcher is already running. Check your system tray.")
        sys.exit(1)

    pythoncom.CoInitialize()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    import signal
    signal.signal(signal.SIGTERM, handle_shutdown)
    signal.signal(signal.SIGINT, handle_shutdown)

    window = BatteryWatcher()
    window.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Application exiting...")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
