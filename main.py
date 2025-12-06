import sys
import os
import subprocess
import json
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import pyqtSignal, QObject
import win32gui
import win32con
import win32process
import keyboard

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def is_window_on_taskbar(hwnd):
    if not win32gui.IsWindowVisible(hwnd):
        return False
    if win32gui.GetParent(hwnd):
        return False
    if win32gui.GetWindow(hwnd, win32con.GW_OWNER):
        return False
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    if style & win32con.WS_EX_TOOLWINDOW:
        return False
    title = win32gui.GetWindowText(hwnd).strip()
    return bool(title)

def get_taskbar_windows(exclude_title="AOT - AlwaysOnTop"):
    windows = []
    seen_pids = set()
    exclude_titles = [exclude_title, "Settings", "Cài đặt"]
    def enum_handler(hwnd, _):
        if is_window_on_taskbar(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in seen_pids:
                seen_pids.add(pid)
                title = win32gui.GetWindowText(hwnd)
                if title not in exclude_titles:
                    windows.append((hwnd, title))
    win32gui.EnumWindows(enum_handler, None)
    return windows

class HotkeySignals(QObject):
    pin_signal = pyqtSignal()
    unpin_signal = pyqtSignal()

class PinApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.pinned_windows = {}
        self.config_file = os.path.join(os.path.expanduser("~"), "AppData", "Local", "AOT_AlwaysOnTop", "config.json")
        
        self.hotkey_pin = "ctrl+shift+p"
        self.hotkey_unpin = "ctrl+shift+u"
        self.close_to_tray = False
        self.language = "vi"
        self.load_config_all()
        
        self.settings_changed = False
        
        self.texts = {
            "vi": {
                "title": "AOT - AlwaysOnTop",
                "select_window": "Chọn cửa sổ để ghim",
                "unpin_all": "Bỏ tất cả ghim",
                "hotkey_label": "Phím tắt: Ghim [{0}] | Bỏ ghim [{1}]",
                "menu_file": "File",
                "menu_open_location": "Mở thư mục file",
                "menu_about": "Thông tin",
                "menu_exit": "Thoát",
                "menu_theme": "Giao diện",
                "menu_white": "Trắng",
                "menu_gray": "Xám",
                "menu_black": "Đen",
                "menu_refresh": "Làm mới",
                "menu_settings": "Cài đặt",
                "settings_title": "Cài đặt",
                "settings_hotkey": "Phím tắt",
                "settings_pin": "Phím tắt ghim:",
                "settings_unpin": "Phím tắt bỏ ghim:",
                "settings_example": "Ví dụ: ctrl+shift+p, alt+p\nCác phím: ctrl, shift, alt, win",
                "settings_close_tray": "Ẩn xuống khay khi đóng",
                "settings_language": "Ngôn ngữ",
                "settings_vietnamese": "Tiếng Việt",
                "settings_english": "English",
                "btn_save": "Lưu",
                "btn_cancel": "Hủy",
                "msg_error": "Lỗi",
                "msg_empty_hotkey": "Vui lòng nhập đầy đủ phím tắt!",
                "msg_invalid_hotkey": "Phím tắt không hợp lệ: {0}",
                "msg_cannot_open": "Không thể mở thư mục: {0}",
                "tray_open": "Mở",
                "tray_exit": "Thoát",
                "about_title": "Thông tin phần mềm",
                "about_app_name": "Tên: AOT - AlwaysOnTop",
                "about_version": "Phiên bản: 2.0.0",
                "about_author": "Tác giả: Ky Khanh Nguyen"
            },
            "en": {
                "title": "AOT - AlwaysOnTop",
                "select_window": "Select window to pin",
                "unpin_all": "Unpin all",
                "hotkey_label": "Hotkey: Pin [{0}] | Unpin [{1}]",
                "menu_file": "File",
                "menu_open_location": "Open file location",
                "menu_about": "About",
                "menu_exit": "Exit",
                "menu_theme": "Theme",
                "menu_white": "White",
                "menu_gray": "Gray",
                "menu_black": "Black",
                "menu_refresh": "Refresh",
                "menu_settings": "Settings",
                "settings_title": "Settings",
                "settings_hotkey": "Hotkeys",
                "settings_pin": "Pin hotkey:",
                "settings_unpin": "Unpin hotkey:",
                "settings_example": "Example: ctrl+shift+p, alt+p\nKeys: ctrl, shift, alt, win",
                "settings_close_tray": "Close to tray",
                "settings_language": "Language",
                "settings_vietnamese": "Tiếng Việt",
                "settings_english": "English",
                "btn_save": "Save",
                "btn_cancel": "Cancel",
                "msg_error": "Error",
                "msg_empty_hotkey": "Please enter hotkeys!",
                "msg_invalid_hotkey": "Invalid hotkey: {0}",
                "msg_cannot_open": "Cannot open folder: {0}",
                "tray_open": "Open",
                "tray_exit": "Exit",
                "about_title": "About",
                "about_app_name": "Name: AOT - AlwaysOnTop",
                "about_version": "Version: 2.0.0",
                "about_author": "Author: Ky Khanh Nguyen"
            }
        }
        
        self.setWindowTitle(self.t("title"))
        self.setGeometry(200, 200, 380, 500)
        
        self.hotkey_signals = HotkeySignals()
        self.hotkey_signals.pin_signal.connect(self.pin_active_window)
        self.hotkey_signals.unpin_signal.connect(self.unpin_active_window)
        
        icon_paths = [
            resource_path("PinApp/icon.ico"),
            resource_path("icon.ico"),
            "PinApp/icon.ico",
            "icon.ico"
        ]
        
        app_icon = None
        for icon_path in icon_paths:
            if os.path.exists(icon_path):
                app_icon = QtGui.QIcon(icon_path)
                self.setWindowIcon(app_icon)
                break
        
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        if app_icon:
            self.tray_icon.setIcon(app_icon)
        else:
            self.tray_icon.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        
        self.create_tray_menu()
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

        self.create_ui()
        
        timer = QtCore.QTimer(self)
        timer.timeout.connect(self.refresh_window_list)
        timer.start(5000)

        saved_theme = self.load_theme()
        self.change_theme(saved_theme)

        self.register_hotkeys()

    def t(self, key):
        return self.texts[self.language].get(key, key)

    def create_tray_menu(self):
        tray_menu = QtWidgets.QMenu()
        open_action = tray_menu.addAction(self.t("tray_open"))
        open_action.triggered.connect(self.show_window)
        exit_action = tray_menu.addAction(self.t("tray_exit"))
        exit_action.triggered.connect(self.quit_app)
        self.tray_icon.setContextMenu(tray_menu)

    def create_ui(self):
        main_layout = QtWidgets.QVBoxLayout(self)
        
        self.menubar = QtWidgets.QMenuBar(self)
        self.create_menus()
        main_layout.setMenuBar(self.menubar)

        self.title_label = QtWidgets.QLabel(self.t("select_window"))
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(self.title_label)

        self.hotkey_label = QtWidgets.QLabel(
            self.t("hotkey_label").format(self.hotkey_pin, self.hotkey_unpin)
        )
        self.hotkey_label.setAlignment(QtCore.Qt.AlignCenter)
        self.hotkey_label.setStyleSheet("font-size: 10px; font-style: italic;")
        main_layout.addWidget(self.hotkey_label)

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        main_layout.addWidget(self.list_widget)

        self.unpin_btn = QtWidgets.QPushButton(self.t("unpin_all"))
        self.unpin_btn.clicked.connect(self.unpin_all_windows)
        main_layout.addWidget(self.unpin_btn)

        self.refresh_window_list()

    def create_menus(self):
        self.menubar.clear()
        
        file_menu = self.menubar.addMenu(self.t("menu_file"))
        open_location_action = QtWidgets.QAction(self.t("menu_open_location"), self)
        open_location_action.triggered.connect(self.open_file_location)
        file_menu.addAction(open_location_action)
        
        about_action = QtWidgets.QAction(self.t("menu_about"), self)
        about_action.triggered.connect(self.show_about)
        file_menu.addAction(about_action)
        
        file_menu.addSeparator()
        exit_action = QtWidgets.QAction(self.t("menu_exit"), self)
        exit_action.triggered.connect(self.quit_app)
        file_menu.addAction(exit_action)

        theme_menu = self.menubar.addMenu(self.t("menu_theme"))
        white_action = QtWidgets.QAction(self.t("menu_white"), self)
        white_action.triggered.connect(lambda: self.change_theme("white"))
        theme_menu.addAction(white_action)
        gray_action = QtWidgets.QAction(self.t("menu_gray"), self)
        gray_action.triggered.connect(lambda: self.change_theme("gray"))
        theme_menu.addAction(gray_action)
        black_action = QtWidgets.QAction(self.t("menu_black"), self)
        black_action.triggered.connect(lambda: self.change_theme("black"))
        theme_menu.addAction(black_action)

        settings_menu = self.menubar.addMenu(self.t("menu_settings"))
        settings_action = QtWidgets.QAction(self.t("settings_title"), self)
        settings_action.triggered.connect(self.show_settings)
        settings_menu.addAction(settings_action)

        refresh_action = self.menubar.addAction(self.t("menu_refresh"))
        refresh_action.triggered.connect(self.refresh_window_list)

    def show_about(self):
        about_dialog = QtWidgets.QDialog(self)
        about_dialog.setWindowTitle(self.t("about_title"))
        about_dialog.setFixedSize(350, 150)
        
        layout = QtWidgets.QVBoxLayout(about_dialog)
        
        app_name_label = QtWidgets.QLabel(self.t("about_app_name"))
        app_name_label.setAlignment(QtCore.Qt.AlignCenter)
        app_name_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(app_name_label)
        
        version_label = QtWidgets.QLabel(self.t("about_version"))
        version_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(version_label)
        
        author_label = QtWidgets.QLabel(self.t("about_author"))
        author_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(author_label)
        
        layout.addStretch()
        
        ok_btn = QtWidgets.QPushButton("OK")
        ok_btn.clicked.connect(about_dialog.accept)
        layout.addWidget(ok_btn)
        
        about_dialog.exec_()

    def update_ui_text(self):
        self.setWindowTitle(self.t("title"))
        self.title_label.setText(self.t("select_window"))
        self.hotkey_label.setText(
            self.t("hotkey_label").format(self.hotkey_pin, self.hotkey_unpin)
        )
        self.unpin_btn.setText(self.t("unpin_all"))
        self.create_menus()
        self.create_tray_menu()

    def show_settings(self):
        self.settings_changed = False
        
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(self.t("settings_title"))
        dialog.setFixedSize(450, 350)
        
        layout = QtWidgets.QVBoxLayout(dialog)
        
        hotkey_group = QtWidgets.QGroupBox(self.t("settings_hotkey"))
        hotkey_layout = QtWidgets.QVBoxLayout()
        
        pin_layout = QtWidgets.QHBoxLayout()
        pin_label = QtWidgets.QLabel(self.t("settings_pin"))
        pin_label.setFixedWidth(120)
        self.pin_input = QtWidgets.QLineEdit(self.hotkey_pin)
        self.pin_input.textChanged.connect(self.on_settings_changed)
        pin_layout.addWidget(pin_label)
        pin_layout.addWidget(self.pin_input)
        hotkey_layout.addLayout(pin_layout)
        
        unpin_layout = QtWidgets.QHBoxLayout()
        unpin_label = QtWidgets.QLabel(self.t("settings_unpin"))
        unpin_label.setFixedWidth(120)
        self.unpin_input = QtWidgets.QLineEdit(self.hotkey_unpin)
        self.unpin_input.textChanged.connect(self.on_settings_changed)
        unpin_layout.addWidget(unpin_label)
        unpin_layout.addWidget(self.unpin_input)
        hotkey_layout.addLayout(unpin_layout)
        
        help_text = QtWidgets.QLabel(self.t("settings_example"))
        help_text.setStyleSheet("font-size: 9px; font-style: italic;")
        hotkey_layout.addWidget(help_text)
        
        hotkey_group.setLayout(hotkey_layout)
        layout.addWidget(hotkey_group)
        
        self.close_to_tray_checkbox = QtWidgets.QCheckBox(self.t("settings_close_tray"))
        self.close_to_tray_checkbox.setChecked(self.close_to_tray)
        self.close_to_tray_checkbox.stateChanged.connect(self.on_settings_changed)
        layout.addWidget(self.close_to_tray_checkbox)
        
        lang_group = QtWidgets.QGroupBox(self.t("settings_language"))
        lang_layout = QtWidgets.QVBoxLayout()
        
        self.lang_vi_radio = QtWidgets.QRadioButton(self.t("settings_vietnamese"))
        self.lang_en_radio = QtWidgets.QRadioButton(self.t("settings_english"))
        
        self.lang_vi_radio.toggled.connect(self.on_settings_changed)
        self.lang_en_radio.toggled.connect(self.on_settings_changed)
        
        if self.language == "vi":
            self.lang_vi_radio.setChecked(True)
        else:
            self.lang_en_radio.setChecked(True)
        
        lang_layout.addWidget(self.lang_vi_radio)
        lang_layout.addWidget(self.lang_en_radio)
        
        lang_group.setLayout(lang_layout)
        layout.addWidget(lang_group)
        
        layout.addStretch()
        
        btn_layout = QtWidgets.QHBoxLayout()
        self.save_btn = QtWidgets.QPushButton(self.t("btn_save"))
        cancel_btn = QtWidgets.QPushButton(self.t("btn_cancel"))
        
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(lambda: self.save_settings(dialog))
        cancel_btn.clicked.connect(dialog.reject)
        
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        dialog.exec_()

    def on_settings_changed(self):
        self.settings_changed = True
        if hasattr(self, 'save_btn'):
            self.save_btn.setEnabled(True)

    def save_settings(self, dialog):
        new_pin = self.pin_input.text().strip().lower()
        new_unpin = self.unpin_input.text().strip().lower()
        
        if not new_pin or not new_unpin:
            QtWidgets.QMessageBox.warning(
                self, self.t("msg_error"), self.t("msg_empty_hotkey")
            )
            return
        
        self.unregister_hotkeys()
        
        self.hotkey_pin = new_pin
        self.hotkey_unpin = new_unpin
        self.close_to_tray = self.close_to_tray_checkbox.isChecked()
        
        old_language = self.language
        self.language = "vi" if self.lang_vi_radio.isChecked() else "en"
        
        self.save_config()
        
        try:
            self.register_hotkeys()
            
            if old_language != self.language:
                self.update_ui_text()
            else:
                self.hotkey_label.setText(
                    self.t("hotkey_label").format(self.hotkey_pin, self.hotkey_unpin)
                )
            
            self.save_btn.setEnabled(False)
            self.settings_changed = False
            dialog.accept()
        except Exception as e:
            QtWidgets.QMessageBox.warning(
                self, self.t("msg_error"), self.t("msg_invalid_hotkey").format(str(e))
            )
            self.load_config_all()
            self.register_hotkeys()

    def tray_icon_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            self.show_window()

    def show_window(self):
        self.show()
        self.activateWindow()

    def quit_app(self):
        self.unregister_hotkeys()
        
        for hwnd, pinned in list(self.pinned_windows.items()):
            if pinned:
                try:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                except Exception:
                    pass
        
        self.tray_icon.hide()
        QtWidgets.QApplication.quit()

    def register_hotkeys(self):
        try:
            keyboard.add_hotkey(self.hotkey_pin, lambda: self.hotkey_signals.pin_signal.emit())
            keyboard.add_hotkey(self.hotkey_unpin, lambda: self.hotkey_signals.unpin_signal.emit())
        except Exception:
            pass

    def unregister_hotkeys(self):
        try:
            keyboard.remove_hotkey(self.hotkey_pin)
            keyboard.remove_hotkey(self.hotkey_unpin)
        except Exception:
            pass

    def pin_active_window(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            
            if title and title != self.t("title"):
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.pinned_windows[hwnd] = True
                self.refresh_window_list()
        except Exception:
            pass

    def unpin_active_window(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            
            if hwnd in self.pinned_windows and self.pinned_windows[hwnd]:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.pinned_windows[hwnd] = False
                self.refresh_window_list()
        except Exception:
            pass

    def load_config_all(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.hotkey_pin = config.get("hotkey_pin", "ctrl+shift+p")
                    self.hotkey_unpin = config.get("hotkey_unpin", "ctrl+shift+u")
                    self.close_to_tray = config.get("close_to_tray", False)
                    self.language = config.get("language", "vi")
        except Exception:
            pass

    def save_config(self):
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            config = {
                "theme": getattr(self, 'current_theme', 'white'),
                "hotkey_pin": self.hotkey_pin,
                "hotkey_unpin": self.hotkey_unpin,
                "close_to_tray": self.close_to_tray,
                "language": self.language
            }
            
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception:
            pass

    def open_file_location(self):
        try:
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(__file__)
            
            subprocess.Popen(f'explorer /select,"{exe_path}"')
        except Exception as e:
            QtWidgets.QMessageBox.warning(
                self, self.t("msg_error"), self.t("msg_cannot_open").format(str(e))
            )

    def save_theme(self, theme):
        self.current_theme = theme
        self.save_config()
    
    def load_theme(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    return config.get("theme", "white")
        except Exception:
            pass
        return "white"

    def change_theme(self, theme):
        self.save_theme(theme)
        
        if theme == "white":
            border_color = "#CCCCCC"
            text_color = "#000000"
            bg_color = "#FFFFFF"
            list_bg = "#F5F5F5"
            btn_bg = "#E0E0E0"
            btn_hover = "#D0D0D0"
            menubar_bg = "#E8E8E8"
            menu_bg = "#FFFFFF"
        elif theme == "gray":
            border_color = "#555555"
            text_color = "#FFFFFF"
            bg_color = "#808080"
            list_bg = "#6B6B6B"
            btn_bg = "#6B6B6B"
            btn_hover = "#757575"
            menubar_bg = "#5A5A5A"
            menu_bg = "#6B6B6B"
        elif theme == "black":
            border_color = "#3E3E3E"
            text_color = "#FFFFFF"
            bg_color = "#1E1E1E"
            list_bg = "#2D2D2D"
            btn_bg = "#2D2D2D"
            btn_hover = "#3A3A3A"
            menubar_bg = "#252525"
            menu_bg = "#2D2D2D"
        
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {bg_color};
                color: {text_color};
            }}
            QListWidget {{
                background-color: {list_bg};
                border: 1px solid {border_color};
            }}
            QPushButton {{
                background-color: {btn_bg};
                border: 1px solid {border_color};
                padding: 5px;
                color: {text_color};
            }}
            QPushButton:hover {{
                background-color: {btn_hover};
            }}
            QPushButton:disabled {{
                background-color: {border_color};
                color: #888888;
            }}
            QMenuBar {{
                background-color: {menubar_bg};
                color: {text_color};
            }}
            QMenuBar::item {{
                background-color: transparent;
                padding: 4px 8px;
            }}
            QMenuBar::item:selected {{
                background-color: #0078D4;
                color: #FFFFFF;
            }}
            QMenuBar::item:pressed {{
                background-color: #005A9E;
                color: #FFFFFF;
            }}
            QMenu {{
                background-color: {menu_bg};
                color: {text_color};
                border: 1px solid {border_color};
            }}
            QMenu::item {{
                padding: 5px 20px;
            }}
            QMenu::item:selected {{
                background-color: #0078D4;
                color: #FFFFFF;
            }}
            QGroupBox {{
                border: 1px solid {border_color};
                margin-top: 10px;
                padding-top: 10px;
                color: {text_color};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}
        """)

    def refresh_window_list(self):
        try:
            self.list_widget.itemChanged.disconnect()
        except:
            pass
        
        self.list_widget.clear()
        windows = get_taskbar_windows()

        for hwnd, title in windows:
            item = QtWidgets.QListWidgetItem(title)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            state = QtCore.Qt.Checked if self.pinned_windows.get(hwnd, False) else QtCore.Qt.Unchecked
            item.setCheckState(state)
            item.setData(QtCore.Qt.UserRole, hwnd)
            self.list_widget.addItem(item)

        self.list_widget.itemChanged.connect(self.toggle_pin)

    def toggle_pin(self, item):
        hwnd = item.data(QtCore.Qt.UserRole)
        checked = item.checkState() == QtCore.Qt.Checked

        try:
            if checked:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.pinned_windows[hwnd] = True
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.pinned_windows[hwnd] = False
        except Exception:
            pass

    def unpin_all_windows(self):
        for hwnd, pinned in list(self.pinned_windows.items()):
            if pinned:
                try:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                    self.pinned_windows[hwnd] = False
                except Exception:
                    pass
        self.refresh_window_list()

    def closeEvent(self, event):
        if self.close_to_tray:
            event.ignore()
            self.hide()
        else:
            self.quit_app()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = PinApp()
    win.show()
    sys.exit(app.exec_())
