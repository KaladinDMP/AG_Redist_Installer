import sys
import json
import os
import logging
import appdirs
import ctypes
import requests
import subprocess
import tempfile
import winreg
import time
import threading                                              
import queue
import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QListWidget, QLabel, 
                             QTextEdit, QMessageBox, QListWidgetItem, QTreeWidgetItem,
                             QStyleFactory, QScrollBar, QComboBox, QCheckBox,
                             QMenu, QAction, QDialog, QTreeWidget, QFrame)  
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QObject, pyqtSlot, QMetaType, QSize, QTimer, QPoint, QPropertyAnimation
from PyQt5.QtGui import QIcon, QPalette, QColor, QFont, QDesktopServices, QTextCursor, QPixmap, QPainter

# Embed the redistributables JSON data
EMBEDDED_REDISTS_JSON = """
{
  "redistributables": {
    "x86": [
      {
        "name": "Visual C++ 2005",
        "version": "8.0.61001",
        "online_url": "https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.exe",
        "install_command": "/q"
      },
      {
        "name": "Visual C++ 2008",
        "version": "9.0.30729.6161",
        "online_url": "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe",
        "install_command": "/qb"
      },
      {
        "name": "Visual C++ 2010",
        "version": "10.0.40219.325",
        "online_url": "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2012",
        "version": "11.0.61030.0",
        "online_url": "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2013",
        "version": "12.0.40664.0",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/10912113/5da66ddebb0ad32ebd4b922fd82e8e25/vcredist_x86.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2015-2022",
        "version": "14.36.32532.0",
        "online_url": "https://aka.ms/vs/17/release/vc_redist.x86.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": ".NET Core 3.1 Runtime",
        "version": "3.1",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/3f353d2c-0431-48c5-bdf6-fbbe8f901bb5/542a4af07c1df5136a98a1c2df6f3d62/windowsdesktop-runtime-3.1.32-win-x86.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 6.0 Desktop Runtime",
        "version": "6.0.33",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/8029cdb3-0f5f-4018-bff7-bacd9b9357f8/daf6c8b102a3bdfbbf235cfa0e46f901/windowsdesktop-runtime-6.0.33-win-x86.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 8.0 Desktop Runtime",
        "version": "8.0.8",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/bd1c2e28-44dd-47bb-a55c-aedd1f3e8cc4/0a15fac821e64cf7b8ec6d99e54e0997/windowsdesktop-runtime-8.0.8-win-x86.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 9.0 Desktop Runtime",
        "version": "9.0.0-rc.1",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/ad33dd90-1911-497e-87d9-f3506c17f87d/2c8aec980e150fa37a65b4bb115bfaf0/windowsdesktop-runtime-9.0.0-rc.1.24452.1-win-x86.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": "7-Zip",
        "version": "24.08",
        "online_url": "https://www.7-zip.org/a/7z2408.exe",
        "install_command": "/S"
      },
      {
        "name": "XNA Framework Redistributable 4.0",
        "version": "4.0",
        "online_url": "https://download.microsoft.com/download/A/C/2/AC2C903B-E6E8-42C2-9FD7-BEBAC362A930/xnafx40_redist.msi",
        "install_command": "/quiet /norestart"
      }
    ],
    "x64": [
      {
        "name": "Visual C++ 2005",
        "version": "8.0.61000",
        "online_url": "https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.exe",
        "install_command": "/q"
      },
      {
        "name": "Visual C++ 2008",
        "version": "9.0.30729.6161",
        "online_url": "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe",
        "install_command": "/qb"
      },
      {
        "name": "Visual C++ 2010",
        "version": "10.0.40219.325",
        "online_url": "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2012",
        "version": "11.0.61030.0",
        "online_url": "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2013",
        "version": "12.0.40664.0",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/10912041/cee5d6bca2ddbcd039da727bf4acb48a/vcredist_x64.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": "Visual C++ 2015-2022",
        "version": "14.36.32532.0",
        "online_url": "https://aka.ms/vs/17/release/vc_redist.x64.exe",
        "install_command": "/passive /norestart"
      },
      {
        "name": ".NET Core 3.1 Runtime",
        "version": "3.1",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/b92958c6-ae36-4efa-aafe-569fced953a5/1654639ef3b20eb576174c1cc200f33a/windowsdesktop-runtime-3.1.32-win-x64.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET Framework 4.7.2",
        "version": "4.7.2",
        "online_url": "https://go.microsoft.com/fwlink/?linkid=863265",
        "install_command": "/install /quiet"
      },
      {
        "name": ".NET Framework 4.8",
        "version": "4.8",
        "online_url": "https://go.microsoft.com/fwlink/?LinkId=2085155",
        "install_command": "/install /quiet"
      },
      {
        "name": ".NET 5.0 Runtime",
        "version": "5.0",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/a0832b5a-6900-442b-af79-6ffddddd6ba4/e2df0b25dd851ee0b38a86947dd0e42e/dotnet-runtime-5.0.17-win-x64.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 6.0 Desktop Runtime",
        "version": "6.0.33",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/3ebc1f91-a5ba-477e-9353-198fa4e13371/35f447d6820b078fd18523764a4f0213/windowsdesktop-runtime-6.0.33-win-x64.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 8.0 Desktop Runtime",
        "version": "8.0.8",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/907765b0-2bf8-494e-93aa-5ef9553c5d68/a9308dc010617e6716c0e6abd53b05ce/windowsdesktop-runtime-8.0.8-win-x64.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET 9.0 Desktop Runtime",
        "version": "9.0.0-rc.1",
        "online_url": "https://download.visualstudio.microsoft.com/download/pr/42d0d927-a9fd-4466-85b9-a92881127771/ada1c6949c9e4a173284391d91add261/windowsdesktop-runtime-9.0.0-rc.1.24452.1-win-x64.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": "7-Zip",
        "version": "24.08",
        "online_url": "https://www.7-zip.org/a/7z2408-x64.exe",
        "install_command": "/S"
      },
      {
        "name": "XNA Framework Redistributable 4.0",
        "version": "4.0",
        "online_url": "https://download.microsoft.com/download/A/C/2/AC2C903B-E6E8-42C2-9FD7-BEBAC362A930/xnafx40_redist.msi",
        "install_command": "/quiet /norestart"
      }
    ],
    "all": [
      {
        "name": "DirectX End-User Runtime",
        "version": "9.29.1974.1",
        "online_url": "https://download.microsoft.com/download/1/7/1/1718CCC4-6315-4D8E-9543-8E28A4E18C4C/dxwebsetup.exe",
        "install_command": "/Q"
      },
      {
        "name": "OpenAL",
        "version": "1.1",
        "online_url": "https://www.dropbox.com/scl/fi/csltc4ueqqwgbp08dsufp/oalinst.exe?rlkey=eyi3z11fova7g85ki3ydkug17&st=3i46c0xi&dl=1",
        "install_command": "/silent"
      },
      {
        "name": ".NET Framework 3.5",
        "version": "3.5",
        "online_url": "https://download.microsoft.com/download/2/0/E/20E90413-712F-438C-988E-FDAA79A8AC3D/dotnetfx35.exe",
        "install_command": "/install /quiet /norestart"
      },
      {
        "name": ".NET Framework 4.0",
        "version": "4.0",
        "online_url": "https://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe",
        "install_command": "/install /quiet"
      },
      {
        "name": "Java Runtime Environment",
        "version": "8u391",
        "online_url": "https://javadl.oracle.com/webapps/download/AutoDL?BundleId=249553_4d245f941845490c91360409ecffb3b4",
        "install_command": "/s"
      },
      {
        "name": "PhysX System Software",
        "version": "9.23.1019",
        "online_url": "https://us.download.nvidia.com/Windows/9.23.1019/PhysX_9.23.1019_SystemSoftware.exe",
        "install_command": "-s"
      }
    ]
  }
}
"""

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class InstallWorker(QThread):
    progress = pyqtSignal(dict)
    finished = pyqtSignal(list)

    def __init__(self, installer, redists_to_install):
        super().__init__()
        self.installer = installer
        self.redists_to_install = redists_to_install

    def run(self):
        results = []
        for redist in self.redists_to_install:
            result = self.installer.install_redist(redist)
            self.progress.emit(result)
            results.append(result)
        self.finished.emit(results)

class ProgressHandler(QObject):
    update_signal = pyqtSignal(dict)

    @pyqtSlot(dict)
    def update_progress(self, result):
        self.update_signal.emit(result)
        
class RedistInstallerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ARMGDDN Games Redist Installer")
        self.setGeometry(100, 100, 800, 750)  # Updated window size

        self.version = "1.0.0"
        self.app_name = "AG Redist Installer"
        self.settings_dir = appdirs.user_data_dir(self.app_name)
        self.settings_file = os.path.join(self.settings_dir, "settings.json")

        self.setup_logging()
        self.load_settings()
        self.load_redists_data()
        self.progress_handler = ProgressHandler()
        self.progress_handler.update_signal.connect(self.update_progress_text)
        self.installations_performed = False
        self.dark_mode = True
        self.installation_status = self.load_installation_status()
        self.init_ui()
        self.create_menu()
        self.style_menus()
        self.apply_style() 
        self.click_count = 0
        self.debug_mode = False
        
        self.create_logo_button()
        
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "icon.ico"))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.debug_timer = QTimer(self)
        self.debug_timer.setSingleShot(True)
        self.debug_timer.timeout.connect(self.hide_debug_message)    
        self.debug_label = QLabel(self)
        self.debug_label.setAlignment(Qt.AlignCenter)
        self.debug_label.setStyleSheet("""
            QLabel {
                background-color: #2A2A2A;
                color: white;
                border: 2px solid #555555;
                border-radius: 10px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        self.debug_label.hide()
    
    def fetch_github_json(self):
        url = "https://raw.githubusercontent.com/KaladinDMP/AG_Redist_Installer/main/redists.json"
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()  # Raises an HTTPError for bad responses
            return response.json()
        except (requests.RequestException, ValueError) as e:
            self.logger.error(f"Failed to fetch JSON from GitHub: {e}")
            return None    
    
    def check_for_updates(self):
        last_check = self.settings.get('last_update_check')
        today = datetime.date.today().isoformat()
        if last_check != today:
            self.load_redists_data()
            self.settings['last_update_check'] = today
            self.save_settings()
        
    def create_logo_button(self):
        original_size = QSize(1024, 890)  # Original image size
        target_size = QSize(100, 87)  # Desired size (maintains aspect ratio)

        self.logo_button = QPushButton(self)
        self.logo_button.setFixedSize(target_size)
        self.logo_button.clicked.connect(self.on_logo_click)

        # Load and scale the image
        pixmap = QPixmap("armgddn_logo.png")
        scaled_pixmap = pixmap.scaled(target_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        # Set the scaled image as an icon
        self.logo_button.setIcon(QIcon(scaled_pixmap))
        self.logo_button.setIconSize(scaled_pixmap.size())

        # Make the button background transparent
        self.logo_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
        """)

        # Position the button in the bottom right corner
        self.position_logo_button()

        # Ensure the button stays in the bottom right corner when window is resized
        self.resizeEvent = self.on_resize

    def position_logo_button(self):
        if hasattr(self, 'logo_button'):
            self.logo_button.move(
                self.width() - self.logo_button.width() - 30,  # Moved 10 pixels to the left
                self.height() - self.logo_button.height() - 30  # Moved 10 pixels up
            )

    def on_resize(self, event):
        # Call the parent class's resizeEvent
        super().resizeEvent(event)
        
        # Reposition the logo button
        self.position_logo_button()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.position_logo_button()
        if self.debug_label.isVisible():
            self.debug_label.move(
                self.width() - self.debug_label.width() - 30,  # Moved 10 pixels to the left
                self.height() - self.debug_label.height() - 90  # Moved above the logo button
            )
            
    def on_logo_click(self):
        self.click_count += 1
        if 5 <= self.click_count < 10:
            clicks_left = 10 - self.click_count
            message = f"{clicks_left} More {'Click' if clicks_left == 1 else 'Clicks'} To Enable Developer Mode"
            self.show_debug_message(message)
        elif self.click_count == 10:
            self.open_debug_window()
            self.click_count = 0
            self.debug_label.hide()

    def show_debug_message(self, message):
        self.debug_label.setText(message)
        self.debug_label.adjustSize()
        self.debug_label.move(
            self.width() - self.debug_label.width() - 30,
            self.height() - self.debug_label.height() - 90
        )
        self.debug_label.show()
        
        # Cancel any existing timer
        self.debug_timer.stop()
        
        # Start a new timer for 2 seconds
        self.debug_timer.start(2000)

    def hide_debug_message(self):
        self.debug_label.hide()

    def open_debug_window(self):
        self.debug_window = QDialog(self)
        self.debug_window.setWindowTitle("Developer Mode")
        self.debug_window.setFixedSize(500, 300)  # Increased height to accommodate new button
        layout = QVBoxLayout()
        
        warning_text = (
            "<p style='font-weight: bold; text-decoration: underline;'>"
            "You Just Activated Developer Mode.<br>"
            "Your Warranty Will Cease To Exist From Here On Out.<br>"
            "Good Luck. You Are Going To Need It.<br>"
            "More Info <a href='https://www.youtube.com/watch?v=dQw4w9WgXcQ' style='color: blue;'>Here</a>"
            "</p>"
        )
        
        warning_label = QLabel(warning_text)
        warning_label.setWordWrap(True)
        warning_label.setAlignment(Qt.AlignCenter)
        warning_label.setTextFormat(Qt.RichText)
        warning_label.setOpenExternalLinks(True)
        font = QFont()
        font.setPointSize(12)
        warning_label.setFont(font)
        
        layout.addWidget(warning_label)
        
        reset_button = QPushButton("Reset Installation Status")
        reset_button.clicked.connect(self.reset_installation_status)
        layout.addWidget(reset_button)
        
        redownload_button = QPushButton("Redownload Redist JSON from GitHub")
        redownload_button.clicked.connect(self.redownload_redist_json)
        layout.addWidget(redownload_button)
        
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.debug_window.close)
        layout.addWidget(close_button)
        
        self.debug_window.setLayout(layout)
        self.debug_window.exec_()

    def redownload_redist_json(self):
            try:
                github_data = self.fetch_github_json()
                if github_data:
                    self.data = github_data
                    self.save_github_json(github_data)
                    self.populate_redist_list()  # Refresh the list with new data
                    QMessageBox.information(self, "Success", "Redistributables JSON successfully updated from GitHub.")
                    self.logger.info("Redistributables JSON successfully updated from GitHub")
                else:
                    QMessageBox.warning(self, "Error", "Failed to fetch redistributables data from GitHub.")
                    self.logger.error("Failed to fetch redistributables data from GitHub")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while updating: {str(e)}")
                self.logger.error(f"Error updating redistributables JSON: {str(e)}", exc_info=True)
    
    def save_github_json(self, data):
            try:
                json_file = os.path.join(self.settings_dir, "github_redists.json")
                with open(json_file, 'w') as f:
                    json.dump(data, f, indent=2)
                self.logger.info(f"GitHub JSON saved to {json_file}")
            except Exception as e:
                self.logger.error(f"Error saving GitHub JSON: {str(e)}", exc_info=True)    
                
    def toggle_developer_mode(self):
        self.developer_mode = not self.developer_mode
        QMessageBox.information(self, "Developer Mode", f"Developer mode {'enabled' if self.developer_mode else 'disabled'}")   
    
    def setup_logging(self):
            log_dir = appdirs.user_log_dir(self.app_name)
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, 'redist_installer.log')
            
            self.logger = logging.getLogger('RedistInstaller')
            self.logger.setLevel(logging.DEBUG)
            fh = logging.FileHandler(log_file)
            fh.setLevel(logging.DEBUG)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)

    def save_settings(self):
        settings = {
            'dark_mode': self.dark_mode,
            'architecture': self.arch,
            'last_update_check': getattr(self, 'last_update_check', None)
        }
        os.makedirs(self.settings_dir, exist_ok=True)
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f)

    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
            else:
                settings = {}
            self.dark_mode = settings.get('dark_mode', True)
            self.arch = settings.get('architecture', 'x64')
            self.last_update_check = settings.get('last_update_check')
        except Exception as e:
            self.logger.error(f"Error loading settings: {str(e)}")
            self.dark_mode = True
            self.arch = 'x64'
            self.last_update_check = 'None'
            
    def save_installation_status(self):
        status_data = {}
        for i in range(self.redist_list.topLevelItemCount()):
            item = self.redist_list.topLevelItem(i)
            redist = item.data(0, Qt.UserRole)
            is_installed = item.data(0, Qt.UserRole + 1)
            is_overridden = item.data(0, Qt.UserRole + 2)
            status_data[f"{redist['name']}_{redist['version']}"] = {
                "installed": is_installed,
                "overridden": is_overridden
            }
        
        status_file = os.path.join(self.settings_dir, "installation_status.json")
        with open(status_file, 'w') as f:
            json.dump(status_data, f)

    def load_installation_status(self):
        status_file = os.path.join(self.settings_dir, "installation_status.json")
        if os.path.exists(status_file):
            with open(status_file, 'r') as f:
                return json.load(f)
        return {}
    
    def load_redists_data(self):
        github_data = self.fetch_github_json()
        if github_data:
            self.logger.info("Successfully loaded redistributables data from GitHub")
            self.data = github_data
        else:
            self.logger.warning("Using embedded redistributables data")
            try:
                self.data = json.loads(EMBEDDED_REDISTS_JSON)
            except json.JSONDecodeError:
                self.logger.error("Failed to load embedded redists data")
                QMessageBox.critical(self, "Error", "Failed to load redistributables data")
                sys.exit(1)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel(f"ARMGDDN Games Redist Installer v{self.version}")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(24)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #4A90E2;")
        layout.addWidget(title_label)

        # Architecture selection
        arch_layout = QHBoxLayout()
        self.arch_label = QLabel("Architecture:")
        self.arch_combo = QComboBox()
        self.arch_combo.addItems(["x64", "x86"])
        self.arch_combo.setCurrentText(self.arch)
        self.arch_combo.currentTextChanged.connect(self.on_arch_changed)
        arch_layout.addWidget(self.arch_label)
        arch_layout.addWidget(self.arch_combo)
        layout.addLayout(arch_layout)

        # Redist List
        self.redist_list = QTreeWidget()
        self.redist_list.setHeaderLabels(["Redistributable", "Status"])
        self.redist_list.setColumnWidth(0, 375)
        self.redist_list.setSelectionMode(QTreeWidget.ExtendedSelection)
        layout.addWidget(self.redist_list)
        
        button_layout = QHBoxLayout()

        # Buttons
        self.install_selected_btn = QPushButton("Install Selected")
        self.install_all_btn = QPushButton("Install All")
        self.install_uninstalled_btn = QPushButton("Install Not Installed")
        self.exit_btn = QPushButton("Exit")

        self.style_buttons()

        self.install_selected_btn.clicked.connect(self.install_selected)
        self.install_all_btn.clicked.connect(self.install_all)
        self.install_uninstalled_btn.clicked.connect(self.install_uninstalled) 
        self.exit_btn.clicked.connect(self.close)

        button_layout.addWidget(self.install_selected_btn)
        button_layout.addWidget(self.install_all_btn)
        button_layout.addWidget(self.install_uninstalled_btn) 
        button_layout.addWidget(self.exit_btn)
        layout.addLayout(button_layout)

       
        # Progress
        self.progress_text = QTextEdit()
        self.progress_text.setReadOnly(True)
        self.progress_text.setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: #2A2A2A;
                width: 14px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #5A5A5A;
                min-height: 20px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
        """)
        layout.addWidget(self.progress_text)

        self.populate_redist_list()

    def style_menus(self):
        menu_style = """
        QMenu {
            background-color: #2E2E2E;
            color: #E0E0E0;
            border: 1px solid #555555;
        }
        QMenu::item {
            padding: 5px 30px 5px 30px;
        }
        QMenu::item:selected {
            background-color: #4A4A4A;
        }
        """
        
        self.menuBar().setStyleSheet(menu_style)
    
    def style_buttons(self):
        for btn in [self.install_selected_btn, self.install_all_btn, self.install_uninstalled_btn, self.exit_btn]:
            if self.dark_mode:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #3A3A3A;
                        color: #E0E0E0;
                        border: none;
                        padding: 5px;
                    }
                    QPushButton:hover {
                        background-color: #4A4A4A;
                    }
                    QPushButton:pressed {
                        background-color: #2A2A2A;
                    }
                """)
            else:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #E0E0E0;
                        color: #3A3A3A;
                        border: none;
                        padding: 5px;
                    }
                    QPushButton:hover {
                        background-color: #F0F0F0;
                    }
                    QPushButton:pressed {
                        background-color: #D0D0D0;
                    }
                """)
                    
    def apply_style(self):
        if self.dark_mode:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #2E2E2E; color: #E0E0E0; }
                QPushButton { 
                    background-color: #4A4A4A; 
                    color: #E0E0E0; 
                    border: none;
                    padding: 5px;
                }
                QPushButton:hover { background-color: #5A5A5A; }
                QPushButton:pressed { background-color: #3A3A3A; }
                QTreeWidget { 
                    background-color: #3E3E3E; 
                    color: #E0E0E0; 
                    alternate-background-color: #353535;
                }
                QTreeWidget::item:selected { 
                    background-color: #4A90E2; 
                    color: #FFFFFF;
                }
                QTreeWidget::item {
                    padding: 5px;
                }
                QTreeWidget::item:hover {
                    background-color: #4A4A4A;
                }
                QHeaderView::section {
                    background-color: #2A2A2A;
                    color: #E0E0E0;
                    padding: 5px;
                    border: 1px solid #3A3A3A;
                }
                QTextEdit { background-color: #3E3E3E; color: #E0E0E0; }
                QComboBox, QCheckBox { 
                    background-color: #3E3E3E;
                    color: #E0E0E0; 
                }
                QComboBox QAbstractItemView {
                    background-color: #3E3E3E;
                    color: #E0E0E0;
                }
                QLabel { color: #E0E0E0; }
                QMenu {
                    background-color: #2E2E2E;
                    color: #E0E0E0;
                    border: 1px solid #555555;
                }
                QMenu::item {
                    padding: 5px 30px 5px 30px;
                }
                QMenu::item:selected {
                    background-color: #4A4A4A;
                }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #FFFFFF; color: #000000; }
                QPushButton { 
                    background-color: #E0E0E0; 
                    color: #000000; 
                    border: none;
                    padding: 5px;
                }
                QPushButton:hover { background-color: #D0D0D0; }
                QPushButton:pressed { background-color: #C0C0C0; }
                QTreeWidget { 
                    background-color: #FFFFFF; 
                    color: #000000; 
                    alternate-background-color: #F5F5F5;
                }
                QTreeWidget::item:selected { 
                    background-color: #4A90E2; 
                    color: #FFFFFF;
                }
                QTreeWidget::item {
                    padding: 5px;
                }
                QTreeWidget::item:hover {
                    background-color: #E0E0E0;
                }
                QHeaderView::section {
                    background-color: #E0E0E0;
                    color: #000000;
                    padding: 5px;
                    border: 1px solid #C0C0C0;
                }
                QTextEdit { background-color: #FFFFFF; color: #000000; }
                QComboBox, QCheckBox { 
                    background-color: #FFFFFF;
                    color: #000000; 
                }
                QComboBox QAbstractItemView {
                    background-color: #FFFFFF;
                    color: #000000;
                }
                QLabel { color: #000000; }
                QMenu {
                    background-color: #FFFFFF;
                    color: #000000;
                    border: 1px solid #C0C0C0;
                }
                QMenu::item {
                    padding: 5px 30px 5px 30px;
                }
                QMenu::item:selected {
                    background-color: #E0E0E0;
                }
            """)
        self.redist_list.setAlternatingRowColors(True)
        self.style_buttons()
        self.redist_list.header().reset()

    def create_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        options_menu = menubar.addMenu('Options')

        info_action = QAction('Info', self)
        info_action.triggered.connect(self.show_info)
        file_menu.addAction(info_action)

        contact_action = QAction('Contact', self)
        contact_action.triggered.connect(self.show_contact)
        file_menu.addAction(contact_action)

        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        dark_mode_action = QAction('Dark Mode', self)
        dark_mode_action.setCheckable(True)
        dark_mode_action.setChecked(self.dark_mode)
        dark_mode_action.triggered.connect(self.toggle_dark_mode)
        options_menu.addAction(dark_mode_action)
        # Add a separator
        options_menu.addSeparator()
        
    def show_info(self):
        info_text = (
            "<div style='font-size: 18px;'>"
            "<p><b><a href='https://t.me/SickSoThr33' style='color: blue; text-decoration: underline;'>DeliciousMeatPop</a> made this tool for "
            "<a href='https://t.me/ARMGDDNGames' style='color: blue; text-decoration: underline;'>ARMGDDN Games</a></b></p>"
            "<p><b>To select multiple items in a row:</b><br>"
            "Hold Shift, then click the first and last item you want to select.</p>"
            "<p><b>To select multiple items not in a row:</b><br>"
            "Hold Ctrl, then click each item you want to select.</p>"
            "</div>"
        )
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Info")
        msg_box.setText(info_text)
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setStyleSheet("QLabel{min-width: 500px;}")
        msg_box.exec_()

    def show_contact(self):
        contact_dialog = QDialog(self)
        contact_dialog.setWindowTitle("Contact")
        layout = QVBoxLayout()

        def create_contact_button(text, url, color):
            button = QPushButton(text)
            button.setStyleSheet(f"""
                QPushButton {{
                    color: {color};
                    font-size: 18px;
                    padding: 15px;
                    min-width: 250px;
                    min-height: 60px;
                }}
            """)
            button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(url)))
            return button

        layout.addWidget(create_contact_button("DMP on Reddit", "https://www.reddit.com/user/DeliciousMeatPop/", "black"))
        layout.addWidget(create_contact_button("DMP on Telegram", "https://t.me/SickSoThr33", "blue"))
        layout.addWidget(create_contact_button("DMP on Discord", "https://discordapp.com/users/191105213808115712", "red"))
        layout.addWidget(create_contact_button("DMP on Github", "https://github.com/KaladinDMP", "green"))

        contact_dialog.setLayout(layout)
        contact_dialog.exec_()

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        self.apply_style()
        self.update_all_item_colors()
        self.style_buttons()
        self.save_settings()  # Add this line
        
    def update_all_item_colors(self):
        for i in range(self.redist_list.topLevelItemCount()):
            item = self.redist_list.topLevelItem(i)
            is_installed = item.data(0, Qt.UserRole + 1)
            self.update_item_color(item, is_installed)

    def on_arch_changed(self, new_arch):
        self.arch = new_arch
        self.save_settings()
        self.populate_redist_list()

    def populate_redist_list(self):
        self.redist_list.clear()
        for redist in self.data['redistributables'].get(self.arch, []) + self.data['redistributables'].get('all', []):
            item = QTreeWidgetItem(self.redist_list)
            item.setText(0, f"{redist['name']} - {redist['version']}")
            item.setData(0, Qt.UserRole, redist)
            
            # Load the saved installation status
            status_key = f"{redist['name']}_{redist['version']}"
            saved_status = self.installation_status.get(status_key, {})
            
            if isinstance(saved_status, bool):
                # Handle old format (just a boolean)
                is_installed = saved_status
                is_overridden = True  # Assume it was overridden in the old format
            else:
                # Handle new format (dictionary)
                is_installed = saved_status.get("installed", False)
                is_overridden = saved_status.get("overridden", False)
            
            if not is_overridden:
                # If not overridden, perform the check
                check_result = self.check_installation_status(redist)
                is_installed = check_result["status"] == "Installed"
            
            item.setData(0, Qt.UserRole + 1, is_installed)
            item.setData(0, Qt.UserRole + 2, is_overridden)
            item.setText(1, "Installed" if is_installed else "Not Installed")
            
            self.redist_list.addTopLevelItem(item)
            self.update_item_color(item, is_installed)

    def update_installation_status(self, result):
        name = result["name"]
        version = result["version"]
        is_installed = result["status"] == "Installed"
        for i in range(self.redist_list.topLevelItemCount()):
            item = self.redist_list.topLevelItem(i)
            item_redist = item.data(0, Qt.UserRole)
            if item_redist["name"] == name and item_redist["version"] == version:
                status = "Installed" if is_installed else "Not Installed"
                item.setText(1, status)
                item.setData(0, Qt.UserRole + 1, is_installed)
                item.setData(0, Qt.UserRole + 2, True)  # Mark as overridden
                self.update_item_color(item, is_installed)
                
                # Update the installation_status dictionary
                status_key = f"{name}_{version}"
                self.installation_status[status_key] = {
                    "installed": is_installed,
                    "overridden": True
                }
                
                break
        self.redist_list.viewport().update()

    def reset_installation_status(self):
        for i in range(self.redist_list.topLevelItemCount()):
            item = self.redist_list.topLevelItem(i)
            redist = item.data(0, Qt.UserRole)
            check_result = self.check_installation_status(redist)
            is_installed = check_result["status"] == "Installed"
            item.setData(0, Qt.UserRole + 1, is_installed)
            item.setData(0, Qt.UserRole + 2, False)  # Reset overridden status
            item.setText(1, "Installed" if is_installed else "Not Installed")
            self.update_item_color(item, is_installed)
        
        self.installation_status = {}  # Clear the saved status
        self.save_installation_status()
        self.redist_list.viewport().update()            
        
    def check_installation_status(self, redist):
        name = redist['name']
        version = redist['version']
        
        # Check registry for installation
        registry_keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        for key in registry_keys:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as reg_key:
                    for i in range(winreg.QueryInfoKey(reg_key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            with winreg.OpenKey(reg_key, subkey_name) as subkey:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                if name.lower() in display_name.lower():
                                    installed_version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                    if self.compare_versions(installed_version, version) >= 0:
                                        return {"name": name, "version": version, "status": "Installed"}
                        except WindowsError:
                            continue
            except WindowsError:
                continue
        
        # Additional checks for specific cases
        if "Visual C++" in name:
            is_installed = self.check_vcredist_status(redist)
        elif "DirectX" in name:
            is_installed = self.check_directx_status()
        elif ".NET" in name:
            is_installed = self.check_dotnet_status(redist)
        else:
            is_installed = False
        
        return {"name": name, "version": version, "status": "Installed" if is_installed else "Not Installed"}
    
    def check_vcredist_status(self, redist):
        # Check common VC++ registry locations
        registry_keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        for key in registry_keys:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as reg_key:
                    for i in range(winreg.QueryInfoKey(reg_key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            with winreg.OpenKey(reg_key, subkey_name) as subkey:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                if redist['name'] in display_name:
                                    version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                    if self.compare_versions(version, redist['version']) >= 0:
                                        return True
                        except WindowsError:
                            continue
            except WindowsError:
                continue
        return False

    def check_directx_status(self):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\DirectX") as key:
                version = winreg.QueryValueEx(key, "Version")[0]
                return self.compare_versions(version, "9.29.1974.1") >= 0
        except WindowsError:
            return False

    def check_dotnet_status(self, redist):
        if "Framework" in redist['name']:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full") as key:
                    release = winreg.QueryValueEx(key, "Release")[0]
                    return release >= 528040  # .NET Framework 4.8
            except WindowsError:
                return False
        else:
            # For .NET Core and .NET 5+, use dotnet --list-runtimes
            try:
                result = subprocess.run(["dotnet", "--list-runtimes"], capture_output=True, text=True, check=True)
                return any(redist['version'] in line for line in result.stdout.split('\n'))
            except subprocess.CalledProcessError:
                return False

    def check_java_status(self, redist):
        registry_key = redist.get('registry_key', '')
        current_version_key = redist.get('current_version_key', 'CurrentVersion')
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_key) as key:
                installed_version = winreg.QueryValueEx(key, current_version_key)[0]
                return self.compare_versions(installed_version, redist['version']) >= 0
        except WindowsError:
            return False

    def compare_versions(self, version1, version2):
        v1_parts = version1.split('.')
        v2_parts = version2.split('.')
        for i in range(max(len(v1_parts), len(v2_parts))):
            v1 = int(v1_parts[i]) if i < len(v1_parts) else 0
            v2 = int(v2_parts[i]) if i < len(v2_parts) else 0
            if v1 > v2:
                return 1
            elif v1 < v2:
                return -1
        return 0

    def update_item_color(self, item, is_installed):
        if self.dark_mode:
            color = QColor('#4CAF50') if is_installed else QColor('#E0E0E0')
        else:
            color = QColor('#4CAF50') if is_installed else QColor('#000000')
        item.setForeground(0, color)
        item.setForeground(1, color)

    def install_selected(self):
            selected_items = self.redist_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "No Selection", "Please select at least one redistributable to install.")
                return
            redists_to_install = [item.data(0, Qt.UserRole) for item in selected_items]
            self.start_installation(redists_to_install)

    def install_all(self):
        redists_to_install = [self.redist_list.topLevelItem(i).data(0, Qt.UserRole) 
                              for i in range(self.redist_list.topLevelItemCount())]
        self.start_installation(redists_to_install)
    
    def install_uninstalled(self):
        uninstalled_redists = []
        for i in range(self.redist_list.topLevelItemCount()):
            item = self.redist_list.topLevelItem(i)
            is_installed = item.data(0, Qt.UserRole + 1)
            if not is_installed:
                redist = item.data(0, Qt.UserRole)
                uninstalled_redists.append(redist)
        
        if not uninstalled_redists:
            QMessageBox.information(self, "No Action Needed", "All redistributables are already installed.")
            return
    
        self.start_installation(uninstalled_redists)

    def start_installation(self, redists_to_install):
        self.worker = InstallWorker(self, redists_to_install)
        self.worker.progress.connect(self.progress_handler.update_progress)
        self.worker.finished.connect(self.show_installation_summary)
        self.worker.start()
    
    @pyqtSlot(dict)
    def update_progress_text(self, result):
        try:
            if result.get('status') == 'No uninstalled redistributables':
                self.progress_text.append("All redistributables are already installed.")
            else:
                message = f"{result['name']} - {result['version']}: {result['status']}"
                self.progress_text.append(message)
            self.progress_text.verticalScrollBar().setValue(self.progress_text.verticalScrollBar().maximum())
            if result['status'] == "Installed" or result['status'].startswith("Not Installed"):
                self.update_installation_status(result)
        except Exception as e:
            self.logger.error(f"Error updating progress text: {str(e)}", exc_info=True)
            self.progress_text.append(f"Error updating progress: {str(e)}")
            
    def install_redist(self, redist):
        name = redist["name"]
        version = redist["version"]
        url = redist["online_url"]
        install_command = redist.get("install_command", "/quiet /norestart")
        
        try:
            self.progress_handler.update_progress({"name": name, "version": version, "status": "Downloading"})
            with tempfile.NamedTemporaryFile(delete=False, suffix='.exe') as temp_file:
                file_path = temp_file.name
                
            response = requests.get(url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            
            with open(file_path, 'wb') as f:
                for data in response.iter_content(8192):
                    size = f.write(data)
                    downloaded = f.tell()
                    percent = int(downloaded / total_size * 100)
                    self.progress_handler.update_progress({"name": name, "version": version, "status": f"Downloading: {percent}%"})
            
            self.progress_handler.update_progress({"name": name, "version": version, "status": "Installing"})
            for attempt in range(3):
                try:
                    result = subprocess.run([file_path] + install_command.split(), 
                                            capture_output=True, text=True, timeout=300)
                    if result.returncode in (0, 3010):
                        self.installations_performed = True
                        return {"name": name, "version": version, "status": "Installed"}
                    elif attempt < 2:
                        self.progress_handler.update_progress({"name": name, "version": version, "status": "Retrying installation"})
                        time.sleep(5)
                    else:
                        return {"name": name, "version": version, "status": "Installation failed"}
                except subprocess.TimeoutExpired:
                    if attempt < 2:
                        self.progress_handler.update_progress({"name": name, "version": version, "status": "Installation timed out, retrying"})
                        time.sleep(5)
                    else:
                        return {"name": name, "version": version, "status": "Installation failed (timeout)"}
        except Exception as e:
            self.logger.error(f"Error installing {name}: {str(e)}", exc_info=True)
            return {"name": name, "version": version, "status": f"Error: {str(e)}"}
        finally:
            if os.path.exists(file_path):
                try:
                    os.unlink(file_path)
                except Exception as e:
                    self.logger.error(f"Error removing temporary file {file_path}: {str(e)}")
        
        return {"name": name, "version": version, "status": "Not Installed"}

    def show_installation_summary(self, results):
        summary = "\n".join([f"{result['name']} - {result['version']}: {result['status']}" for result in results])
        QMessageBox.information(self, "Installation Summary", f"Installation process completed.\n\n{summary}")
        for result in results:
            self.update_installation_status(result)

    def closeEvent(self, event):
        self.debug_mode = False
        self.click_count = 0
        reply = QMessageBox.question(self, 'Exit', 'Are you sure you want to exit?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.save_settings()
            self.save_installation_status()  # Save installation status
            if self.installations_performed:
                restart_reply = QMessageBox.question(self, 'Restart', 'Do you want to restart your computer now?',
                                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                
                if restart_reply == QMessageBox.Yes:
                    try:
                        subprocess.run(["shutdown", "/r", "/t", "10"], check=True)
                    except subprocess.CalledProcessError as e:
                        QMessageBox.warning(self, 'Restart Failed', f'Failed to initiate restart: {e}')
                        self.logger.error(f"Failed to initiate restart: {e}")
                else:
                    QMessageBox.information(self, 'Restart Later', 'Please remember to restart your computer later to complete the installation process.')
            
            event.accept()
        else:
            event.ignore()

if __name__ == '__main__':
    if not is_admin():
        run_as_admin()
    else:
        app = QApplication(sys.argv)
        ex = RedistInstallerGUI()
        ex.show()
        sys.exit(app.exec_())