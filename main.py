import sys
import os
import subprocess
import json
from PyQt5 import QtWidgets, QtGui, QtCore
import win32gui
import win32con
import win32process
from plyer import notification

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
    
    exclude_titles = [
        exclude_title,
        "Settings",
        "Cài đặt"
    ]

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

class PinApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AOT - AlwaysOnTop")
        self.setGeometry(200, 200, 380, 500)

        self.pinned_windows = {}
        
        self.config_file = os.path.join(os.path.expanduser("~"), "AppData", "Local", "AOT_AlwaysOnTop", "config.json")
        
        icon_paths = [
            resource_path("PinApp/icon.ico"),
            resource_path("icon.ico"),
            "PinApp/icon.ico",
            "icon.ico"
        ]
        
        for icon_path in icon_paths:
            if os.path.exists(icon_path):
                self.setWindowIcon(QtGui.QIcon(icon_path))
                break

        main_layout = QtWidgets.QVBoxLayout(self)

        menubar = QtWidgets.QMenuBar(self)
        
        file_menu = menubar.addMenu("File")

        open_location_action = QtWidgets.QAction("Open file location", self)
        open_location_action.triggered.connect(self.open_file_location)
        file_menu.addAction(open_location_action)

        file_menu.addSeparator()

        exit_action = QtWidgets.QAction("Exit", self)
        exit_action.triggered.connect(self.exit_app)
        file_menu.addAction(exit_action)

        theme_menu = menubar.addMenu("Theme")

        white_action = QtWidgets.QAction("Trắng", self)
        white_action.triggered.connect(lambda: self.change_theme("white"))
        theme_menu.addAction(white_action)
        
        gray_action = QtWidgets.QAction("Xám", self)
        gray_action.triggered.connect(lambda: self.change_theme("gray"))
        theme_menu.addAction(gray_action)
        
        black_action = QtWidgets.QAction("Đen", self)
        black_action.triggered.connect(lambda: self.change_theme("black"))
        theme_menu.addAction(black_action)

        refresh_menu = menubar.addMenu("Refresh")
        
        refresh_action = QtWidgets.QAction("Refresh window list", self)
        refresh_action.triggered.connect(self.refresh_window_list)
        refresh_menu.addAction(refresh_action)
        
        main_layout.setMenuBar(menubar)

        title = QtWidgets.QLabel("Chọn cửa sổ để ghim")
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        main_layout.addWidget(self.list_widget)

        self.unpin_btn = QtWidgets.QPushButton("Bỏ tất cả ghim")
        self.unpin_btn.clicked.connect(self.unpin_all_windows)
        main_layout.addWidget(self.unpin_btn)

        self.refresh_window_list()

        timer = QtCore.QTimer(self)
        timer.timeout.connect(self.refresh_window_list)
        timer.start(5000)

        saved_theme = self.load_theme()
        self.change_theme(saved_theme)

    def open_file_location(self):
        """Mở thư mục chứa file exe đang chạy"""
        try:

            if getattr(sys, 'frozen', False):

                exe_path = sys.executable
            else:
 
                exe_path = os.path.abspath(__file__)
            
         
            exe_dir = os.path.dirname(exe_path)
    
            subprocess.Popen(f'explorer /select,"{exe_path}"')
            
        except Exception as e:
            QtWidgets.QMessageBox.warning(
                self,
                "Lỗi",
                f"Không thể mở thư mục: {str(e)}"
            )

    def exit_app(self):
        """Thoát ứng dụng"""
        self.close()

    def save_theme(self, theme):
        """Lưu theme vào file cấu hình"""
        try:

            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
         
            with open(self.config_file, 'w') as f:
                json.dump({"theme": theme}, f)
        except Exception as e:
            print(f"Không thể lưu theme: {e}")
    
    def load_theme(self):
        """Load theme từ file cấu hình"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    return config.get("theme", "white")
        except Exception as e:
            print(f"Không thể load theme: {e}")
        
     
        return "white"

    def change_theme(self, theme):

        self.save_theme(theme)
        
        if theme == "white":
            self.setStyleSheet("""
                QWidget {
                    background-color: #FFFFFF;
                    color: #000000;
                }
                QListWidget {
                    background-color: #F5F5F5;
                    border: 1px solid #CCCCCC;
                }
                QPushButton {
                    background-color: #E0E0E0;
                    border: 1px solid #AAAAAA;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #D0D0D0;
                }
                QMenuBar {
                    background-color: #F0F0F0;
                }
                QMenuBar::item:selected {
                    background-color: #E0E0E0;
                }
                QMenu {
                    background-color: #FFFFFF;
                }
                QMenu::item:selected {
                    background-color: #E0E0E0;
                }
            """)
        elif theme == "gray":
            self.setStyleSheet("""
                QWidget {
                    background-color: #808080;
                    color: #FFFFFF;
                }
                QListWidget {
                    background-color: #6B6B6B;
                    border: 1px solid #555555;
                }
                QPushButton {
                    background-color: #6B6B6B;
                    border: 1px solid #555555;
                    padding: 5px;
                    color: #FFFFFF;
                }
                QPushButton:hover {
                    background-color: #757575;
                }
                QMenuBar {
                    background-color: #6B6B6B;
                    color: #FFFFFF;
                }
                QMenuBar::item:selected {
                    background-color: #757575;
                }
                QMenu {
                    background-color: #6B6B6B;
                    color: #FFFFFF;
                }
                QMenu::item:selected {
                    background-color: #757575;
                }
            """)
        elif theme == "black":
            self.setStyleSheet("""
                QWidget {
                    background-color: #1E1E1E;
                    color: #FFFFFF;
                }
                QListWidget {
                    background-color: #2D2D2D;
                    border: 1px solid #3E3E3E;
                }
                QPushButton {
                    background-color: #2D2D2D;
                    border: 1px solid #3E3E3E;
                    padding: 5px;
                    color: #FFFFFF;
                }
                QPushButton:hover {
                    background-color: #3A3A3A;
                }
                QMenuBar {
                    background-color: #2D2D2D;
                    color: #FFFFFF;
                }
                QMenuBar::item:selected {
                    background-color: #3A3A3A;
                }
                QMenu {
                    background-color: #2D2D2D;
                    color: #FFFFFF;
                }
                QMenu::item:selected {
                    background-color: #3A3A3A;
                }
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
                self.show_notification(item.text(), "đã được ghim lên trên cùng")
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.pinned_windows[hwnd] = False
                self.show_notification(item.text(), "đã bỏ ghim khỏi trên cùng")
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
        """Tự động bỏ ghim tất cả cửa sổ khi đóng app"""
        for hwnd, pinned in list(self.pinned_windows.items()):
            if pinned:
                try:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                except Exception:
                    pass
        event.accept()

    @staticmethod
    def show_notification(title, message):
        notification.notify(
            title="AOT - AlwaysOnTop",
            message=f"{title} - {message}",
            timeout=1
        )

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = PinApp()
    win.show()
    sys.exit(app.exec_())