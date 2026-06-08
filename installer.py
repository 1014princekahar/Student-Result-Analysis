import sys
import os
import shutil
import urllib.request
import zipfile
import subprocess
import tempfile
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QStackedWidget, 
                               QProgressBar, QTextEdit, QFileDialog, QCheckBox)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon, QPixmap

# Theme and Styles
INSTALLER_STYLE = """
QWidget {
    background-color: #0c1a3a;
    color: #f1f5f9;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 13px;
}

QLabel#title {
    color: #ffffff;
    font-size: 22px;
    font-weight: 700;
}

QLabel#subtitle {
    color: #f97316;
    font-size: 14px;
    font-weight: 600;
}

QLabel#info_text {
    color: #94a3b8;
    font-size: 13px;
    line-height: 1.5;
}

QPushButton {
    background-color: #f97316;
    color: white;
    border-radius: 5px;
    padding: 9px 22px;
    font-weight: 700;
    border: none;
    font-size: 13px;
    min-width: 90px;
}
QPushButton:hover {
    background-color: #ea580c;
}
QPushButton:pressed {
    background-color: #c2410c;
}
QPushButton:disabled {
    background-color: #1e293b;
    color: #475569;
}

QPushButton#secondary {
    background-color: transparent;
    border: 1px solid #94a3b8;
    color: #e2e8f0;
}
QPushButton#secondary:hover {
    background-color: rgba(255,255,255,0.05);
}

QProgressBar {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 6px;
    text-align: center;
    color: white;
    font-weight: bold;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #0d9488;
    border-radius: 5px;
}

QTextEdit#log {
    background-color: #020617;
    color: #38bdf8;
    border: 1px solid #1e293b;
    border-radius: 6px;
    font-family: 'Consolas', monospace;
    font-size: 11px;
}

QLineEdit {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 5px;
    padding: 8px;
    color: white;
}
"""

class InstallWorker(QThread):
    progress = Signal(int)
    status = Signal(str)
    log = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, dest_dir, create_shortcut):
        super().__init__()
        self.dest_dir = dest_dir
        self.create_shortcut = create_shortcut
        self.python_path = "python" # fallback

    def run(self):
        try:
            # Step 1: Detect Python path
            self.status.emit("Detecting Python environment...")
            self.progress.emit(5)
            
            # Check if python is accessible
            try:
                result = subprocess.run(["python", "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
                self.log.emit(f"Python detected: {result.stdout.strip() or result.stderr.strip()}")
            except Exception:
                self.log.emit("Python not found in system PATH. Attempting standard paths...")
                # Try standard windows install paths
                local_appdata = os.environ.get("LOCALAPPDATA", "")
                python_local_dir = os.path.join(local_appdata, "Programs", "Python")
                found = False
                if os.path.exists(python_local_dir):
                    for folder in os.listdir(python_local_dir):
                        path = os.path.join(python_local_dir, folder, "python.exe")
                        if os.path.exists(path):
                            self.python_path = path
                            self.log.emit(f"Found Python at: {self.python_path}")
                            found = True
                            break
                if not found:
                    self.finished.emit(False, "Python is not installed. Please install Python first.")
                    return

            # Step 2: Create destination folder
            self.status.emit("Preparing installation directory...")
            self.progress.emit(10)
            os.makedirs(self.dest_dir, exist_ok=True)
            self.log.emit(f"Installation directory: {self.dest_dir}")

            # Step 3: Download Zip
            self.status.emit("Downloading application files from GitHub...")
            self.progress.emit(20)
            zip_url = "https://github.com/1014princekahar/Student-Result-Analysis/archive/refs/heads/main.zip"
            temp_zip = os.path.join(tempfile.gettempdir(), "student_analyzer.zip")
            
            self.log.emit(f"Downloading ZIP from {zip_url}...")
            
            def download_progress(count, block_size, total_size):
                percent = int(count * block_size * 100 / total_size)
                # scale to 20% - 50%
                scaled = 20 + int(percent * 0.3)
                self.progress.emit(min(scaled, 50))
            
            urllib.request.urlretrieve(zip_url, temp_zip, download_progress)
            self.log.emit("Download complete.")

            # Step 4: Extract files
            self.status.emit("Extracting application files...")
            self.progress.emit(55)
            self.log.emit("Extracting files to temporary folder...")
            
            temp_extract = os.path.join(tempfile.gettempdir(), "student_analyzer_extract")
            if os.path.exists(temp_extract):
                shutil.rmtree(temp_extract)
            os.makedirs(temp_extract)

            with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
                zip_ref.extractall(temp_extract)
            
            # Clean up zip
            os.remove(temp_zip)

            # Move contents from extracted subfolder to destination
            subfolders = os.listdir(temp_extract)
            if subfolders:
                extracted_root = os.path.join(temp_extract, subfolders[0])
                self.log.emit("Copying files to installation folder...")
                for item in os.listdir(extracted_root):
                    s = os.path.join(extracted_root, item)
                    d = os.path.join(self.dest_dir, item)
                    if os.path.exists(d):
                        if os.path.isdir(d):
                            shutil.rmtree(d)
                        else:
                            os.remove(d)
                    if os.path.isdir(s):
                        shutil.copytree(s, d)
                    else:
                        shutil.copy2(s, d)
            
            # Clean up extract temp dir
            shutil.rmtree(temp_extract)
            self.log.emit("Extraction and file copy complete.")

            # Step 5: Setup Virtual Environment
            self.status.emit("Creating Python Virtual Environment (venv)...")
            self.progress.emit(65)
            self.log.emit("Initializing virtual environment...")
            venv_dir = os.path.join(self.dest_dir, "venv")
            
            cmd_venv = [self.python_path, "-m", "venv", "venv"]
            res = subprocess.run(cmd_venv, cwd=self.dest_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if res.returncode != 0:
                self.log.emit(f"Error creating venv: {res.stderr}")
                self.finished.emit(False, "Failed to create Virtual Environment.")
                return
            self.log.emit("Virtual environment created.")

            # Step 6: Install dependencies
            self.status.emit("Installing application dependencies (PySide6, pandas, etc.)...")
            self.progress.emit(75)
            self.log.emit("Installing requirements. This may take a moment...")
            
            pip_path = os.path.join(venv_dir, "Scripts", "pip.exe")
            cmd_pip = [pip_path, "install", "-r", "requirements.txt"]
            
            process = subprocess.Popen(cmd_pip, cwd=self.dest_dir, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    clean_out = output.strip()
                    if clean_out:
                        self.log.emit(clean_out)
            
            rc = process.poll()
            if rc != 0:
                self.finished.emit(False, "Failed to install Python dependencies.")
                return
            
            self.log.emit("Dependencies installed successfully.")

            # Step 7: Create shortcuts if selected
            if self.create_shortcut:
                self.status.emit("Creating Desktop Shortcut...")
                self.progress.emit(95)
                self.create_desktop_shortcut()

            self.progress.emit(100)
            self.status.emit("Installation Completed Successfully!")
            self.finished.emit(True, "Success")
        except Exception as e:
            self.log.emit(f"Exception error: {str(e)}")
            self.finished.emit(False, str(e))

    def create_desktop_shortcut(self):
        try:
            desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            shortcut_path = os.path.normpath(os.path.join(desktop, "Student Result Analyzer.lnk"))
            target = os.path.normpath(os.path.join(self.dest_dir, "venv", "Scripts", "pythonw.exe"))
            args = f'"{os.path.normpath(os.path.join(self.dest_dir, "main.py"))}"'
            working_dir = os.path.normpath(self.dest_dir)
            icon = os.path.normpath(os.path.join(self.dest_dir, "assets", "icon.ico"))
            
            # PowerShell script to create shortcut with icon index 0 (single backslashes)
            # We escape double quotes in args with a backtick (`) for PowerShell
            args_escaped = args.replace('"', '`"')
            ps_script = (
                f"$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{shortcut_path}'); "
                f"$s.TargetPath = '{target}'; "
                f"$s.Arguments = '{args_escaped}'; "
                f"$s.WorkingDirectory = '{working_dir}'; "
                f"$s.IconLocation = '{icon},0'; "
                f"$s.Save()"
            )
            subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.log.emit("Desktop shortcut created with custom icon.")
        except Exception as e:
            self.log.emit(f"Failed to create desktop shortcut: {e}")


class InstallerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎓 Student Result Analyzer - Setup Wizard")
        self.setFixedSize(550, 400)
        self.setStyleSheet(INSTALLER_STYLE)
        
        # Center Stacked Widget
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        
        # Setup Default Paths
        user_home = os.environ.get("USERPROFILE", "C:\\")
        self.default_dest = os.path.join(user_home, "Student Result Analyzer")

        # Initialize Screens
        self.init_welcome_screen()
        self.init_check_screen()
        self.init_location_screen()
        self.init_progress_screen()
        self.init_success_screen()

        # Try to set app icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

    def init_welcome_screen(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        title = QLabel("Student Result Analyzer")
        title.setObjectName("title")
        
        subtitle = QLabel("Setup Wizard")
        subtitle.setObjectName("subtitle")
        
        desc = QLabel(
            "This setup wizard will download and install the Student Result Analyzer "
            "application on your computer.\n\n"
            "The installation process will:\n"
            " • Verify/Install Python system environment\n"
            " • Download files from GitHub repository\n"
            " • Configure the local database and virtual environment\n"
            " • Create desktop and start menu shortcuts"
        )
        desc.setObjectName("info_text")
        desc.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(desc)
        layout.addStretch()

        # Nav Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        next_btn = QPushButton("Next >")
        next_btn.clicked.connect(self.go_to_check)
        btn_layout.addWidget(next_btn)
        layout.addLayout(btn_layout)

        self.stack.addWidget(page)

    def init_check_screen(self):
        self.check_page = QWidget()
        layout = QVBoxLayout(self.check_page)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        title = QLabel("System Prerequisites Check")
        title.setObjectName("title")
        
        self.python_status = QLabel("Checking Python installation...")
        self.python_status.setObjectName("subtitle")
        
        self.python_desc = QLabel("Please wait while setup searches for Python...")
        self.python_desc.setObjectName("info_text")
        self.python_desc.setWordWrap(True)

        # Action Buttons for Python installation
        self.py_install_btn = QPushButton("Install Python (Recommended)")
        self.py_install_btn.clicked.connect(self.download_and_install_python)
        self.py_install_btn.setVisible(False)

        layout.addWidget(title)
        layout.addWidget(self.python_status)
        layout.addWidget(self.python_desc)
        layout.addWidget(self.py_install_btn)
        layout.addStretch()

        # Nav Buttons
        self.check_btn_layout = QHBoxLayout()
        back_btn = QPushButton("< Back")
        back_btn.setObjectName("secondary")
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.check_next_btn = QPushButton("Next >")
        self.check_next_btn.setEnabled(False)
        self.check_next_btn.clicked.connect(lambda: self.stack.setCurrentIndex(2))
        
        self.check_btn_layout.addWidget(back_btn)
        self.check_btn_layout.addStretch()
        self.check_btn_layout.addWidget(self.check_next_btn)
        layout.addLayout(self.check_btn_layout)

        self.stack.addWidget(self.check_page)

    def init_location_screen(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        title = QLabel("Choose Install Location")
        title.setObjectName("title")
        
        desc = QLabel("Select the folder where you want to install Student Result Analyzer:")
        desc.setObjectName("info_text")
        
        loc_layout = QHBoxLayout()
        self.path_lbl = QLabel(self.default_dest)
        self.path_lbl.setStyleSheet("background-color: #1e293b; padding: 8px; border: 1px solid #334155; border-radius: 5px;")
        browse_btn = QPushButton("Browse...")
        browse_btn.setObjectName("secondary")
        browse_btn.clicked.connect(self.browse_location)
        loc_layout.addWidget(self.path_lbl, 1)
        loc_layout.addWidget(browse_btn)

        self.shortcut_cb = QCheckBox("Create Desktop Shortcut")
        self.shortcut_cb.setChecked(True)
        self.shortcut_cb.setStyleSheet("QCheckBox { spacing: 8px; color: #f1f5f9; }")

        layout.addWidget(title)
        layout.addWidget(desc)
        layout.addLayout(loc_layout)
        layout.addWidget(self.shortcut_cb)
        layout.addStretch()

        # Nav Buttons
        btn_layout = QHBoxLayout()
        back_btn = QPushButton("< Back")
        back_btn.setObjectName("secondary")
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        install_btn = QPushButton("Install")
        install_btn.clicked.connect(self.start_installation)
        
        btn_layout.addWidget(back_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(install_btn)
        layout.addLayout(btn_layout)

        self.stack.addWidget(page)

    def init_progress_screen(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(10)

        title = QLabel("Installing Application")
        title.setObjectName("title")
        
        self.status_lbl = QLabel("Starting installation...")
        self.status_lbl.setObjectName("subtitle")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        self.log_text = QTextEdit()
        self.log_text.setObjectName("log")
        self.log_text.setReadOnly(True)

        layout.addWidget(title)
        layout.addWidget(self.status_lbl)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_text)

        self.stack.addWidget(page)

    def init_success_screen(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        title = QLabel("Installation Complete")
        title.setObjectName("title")
        
        subtitle = QLabel("Finished successfully!")
        subtitle.setObjectName("subtitle")
        
        desc = QLabel(
            "Student Result Analyzer has been successfully installed on your computer.\n\n"
            "You can launch the application now or use the Desktop shortcut created for you."
        )
        desc.setObjectName("info_text")

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(desc)
        layout.addStretch()

        # Nav Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.setObjectName("secondary")
        close_btn.clicked.connect(self.close)
        
        launch_btn = QPushButton("Launch Application")
        launch_btn.clicked.connect(self.launch_application_now)
        
        btn_layout.addWidget(close_btn)
        btn_layout.addWidget(launch_btn)
        layout.addLayout(btn_layout)

        self.stack.addWidget(page)

    # --- Actions & Logic ---
    def go_to_check(self):
        self.stack.setCurrentIndex(1)
        self.check_python()

    def check_python(self):
        self.py_install_btn.setVisible(False)
        self.check_next_btn.setEnabled(False)
        
        # Check standard system python
        try:
            result = subprocess.run(["python", "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
            version = result.stdout.strip() or result.stderr.strip()
            self.python_status.setText("✅ Python detected!")
            self.python_desc.setText(f"System has {version} installed. You are ready to proceed.")
            self.check_next_btn.setEnabled(True)
            return
        except Exception:
            pass

        # Try to search in standard windows AppData directory
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        python_local_dir = os.path.join(local_appdata, "Programs", "Python")
        if os.path.exists(python_local_dir):
            for folder in os.listdir(python_local_dir):
                path = os.path.join(python_local_dir, folder, "python.exe")
                if os.path.exists(path):
                    self.python_status.setText("✅ Python detected!")
                    self.python_desc.setText(f"Found Python at local folder: {folder}. You are ready to proceed.")
                    self.check_next_btn.setEnabled(True)
                    return

        # If not found at all
        self.python_status.setText("❌ Python is not installed")
        self.python_status.setStyleSheet("color: #ef4444; font-weight: bold;")
        self.python_desc.setText(
            "This application requires Python 3.10 or higher. "
            "Please click the button below to install Python automatically, "
            "or download it manually from python.org."
        )
        self.py_install_btn.setVisible(True)

    def download_and_install_python(self):
        self.py_install_btn.setEnabled(False)
        self.python_status.setText("Downloading Python Installer...")
        self.python_status.setStyleSheet("color: #f97316; font-weight: bold;")
        self.python_desc.setText("Please wait while Python installer is downloading from python.org...")
        QApplication.processEvents()
        
        try:
            python_url = "https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe"
            installer_path = os.path.join(tempfile.gettempdir(), "python_setup.exe")
            
            urllib.request.urlretrieve(python_url, installer_path)
            
            self.python_status.setText("Running Python Installer...")
            self.python_desc.setText(
                "Python installer is running. Please complete the setup wizard.\n\n"
                "⚠️ IMPORTANT: Make sure to check the box '[x] Add Python.exe to PATH' at the bottom of the installer window!"
            )
            QApplication.processEvents()

            # Run Python installer in passive mode (shows basic progress bar, requires no input, adds to path)
            process = subprocess.run([installer_path, "/passive", "PrependPath=1"], check=True)
            
            # Clean installer
            try:
                os.remove(installer_path)
            except Exception:
                pass
            
            self.python_status.setText("Re-checking Python...")
            self.python_status.setStyleSheet("color: #f1f5f9; font-weight: bold;")
            self.check_python()
        except Exception as e:
            self.python_status.setText("❌ Python Installation Failed")
            self.python_status.setStyleSheet("color: #ef4444; font-weight: bold;")
            self.python_desc.setText(f"Error: {e}\n\nPlease install Python manually from: https://www.python.org/downloads/")
            self.py_install_btn.setEnabled(True)

    def browse_location(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Install Directory", self.default_dest)
        if dir_path:
            self.default_dest = os.path.normpath(dir_path)
            self.path_lbl.setText(self.default_dest)

    def start_installation(self):
        self.stack.setCurrentIndex(3) # Show progress screen
        self.worker = InstallWorker(self.default_dest, self.shortcut_cb.isChecked())
        
        # Connect signals
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_lbl.setText)
        self.worker.log.connect(self.log_text.append)
        self.worker.finished.connect(self.installation_finished)
        
        self.worker.start()

    def installation_finished(self, success, message):
        if success:
            self.stack.setCurrentIndex(4) # Show success screen
        else:
            self.status_lbl.setText("❌ Installation Failed")
            self.status_lbl.setStyleSheet("color: #ef4444; font-weight: bold;")
            self.log_text.append(f"\n[FATAL ERROR] {message}")
            
            # Add a back button to retry or exit
            btn_layout = QHBoxLayout()
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(self.close)
            btn_layout.addStretch()
            btn_layout.addWidget(close_btn)
            self.centralWidget().layout().addLayout(btn_layout)

    def launch_application_now(self):
        # Start the application
        app_path = os.path.normpath(os.path.join(self.default_dest, "venv", "Scripts", "pythonw.exe"))
        script_path = os.path.normpath(os.path.join(self.default_dest, "main.py"))
        
        if os.path.exists(app_path) and os.path.exists(script_path):
            subprocess.Popen([app_path, script_path], cwd=self.default_dest)
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InstallerApp()
    window.show()
    sys.exit(app.exec())
