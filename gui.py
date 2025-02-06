import sys
import json
import time
import requests
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QFileDialog, QAbstractItemView,
    QLabel, QSpacerItem, QSizePolicy, QMessageBox
)
from PyQt5.QtCore import Qt, QThread, QObject, pyqtSignal, pyqtSlot, QSettings
from PyQt5.QtGui import QFont, QColor, QIcon

UPLOAD_TIMEOUT = 300
POLL_INTERVAL = 10
MAX_RETRIES = 5

# Modern color palette
COLORS = {
    "primary": "#2c3e50",
    "secondary": "#3498db",
    "success": "#27ae60",
    "danger": "#e74c3c",
    "background": "#ecf0f1",
    "text": "#2c3e50"
}

STYLESHEET = f"""
    QMainWindow {{
        background-color: {COLORS['background']};
    }}
    QPushButton {{
        background-color: {COLORS['primary']};
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {COLORS['secondary']};
    }}
    QPushButton:pressed {{
        background-color: {COLORS['success']};
    }}
    QListWidget {{
        background-color: white;
        border: 2px solid {COLORS['primary']};
        border-radius: 4px;
    }}
    QTableWidget {{
        background-color: white;
        border: 2px solid {COLORS['primary']};
        border-radius: 4px;
    }}
    QHeaderView::section {{
        background-color: {COLORS['primary']};
        color: white;
        padding: 6px;
    }}
    QProgressBar {{
        border: 1px solid {COLORS['primary']};
        border-radius: 4px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background-color: {COLORS['secondary']};
        width: 10px;
    }}
"""

UPLOAD_TIMEOUT = 300
POLL_INTERVAL = 10
MAX_RETRIES = 5

class ConversionWorker(QObject):
    progress_updated = pyqtSignal(str, int)
    finished = pyqtSignal(str, bool, str)
    stopped = pyqtSignal()

    def __init__(self, file_path, credentials):
        super().__init__()
        self.file_path = file_path
        self.credentials = credentials
        self._is_running = False
        self.token_cache = None

    def get_access_token(self):
        if self.token_cache and self.token_cache["expires_at"] > time.time():
            return self.token_cache["access_token"]

        url = "https://ims-na1.adobelogin.com/ims/token/v3"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.credentials["client_credentials"]["client_id"],
            "client_secret": self.credentials["client_credentials"]["client_secret"],
            "scope": "openid,AdobeID,read_organizations,exportpdf"
        }
        
        response = requests.post(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=30)
        response.raise_for_status()
        token_data = response.json()
        
        self.token_cache = {
            "access_token": token_data["access_token"],
            "expires_at": time.time() + token_data["expires_in"] - 30
        }
        return self.token_cache["access_token"]

    def run(self):
        self._is_running = True
        try:
            output_path = str(Path(self.file_path).with_suffix(".docx"))
            
            self.progress_updated.emit(self.file_path, 10)
            access_token = self.get_access_token()
            
            self.progress_updated.emit(self.file_path, 30)
            asset_id = self.upload_pdf(access_token)
            
            self.progress_updated.emit(self.file_path, 50)
            job_id = self.convert_pdf_to_docx(access_token, asset_id)
            
            self.progress_updated.emit(self.file_path, 70)
            self.poll_and_download_result(access_token, job_id, output_path)
            
            self.finished.emit(self.file_path, True, "")
            self.progress_updated.emit(self.file_path, 100)

        except Exception as e:
            self.finished.emit(self.file_path, False, str(e))
        finally:
            self.stopped.emit()

    def upload_pdf(self, access_token):
        url = "https://pdf-services.adobe.io/assets"
        headers = {
            "x-api-key": self.credentials["client_credentials"]["client_id"],
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        response = requests.post(url, json={"mediaType": "application/pdf"}, headers=headers, timeout=30)
        response.raise_for_status()
        upload_data = response.json()
        
        with open(self.file_path, "rb") as f:
            requests.put(upload_data["uploadUri"], data=f, headers={"Content-Type": "application/pdf"}, timeout=UPLOAD_TIMEOUT)
        
        return upload_data["assetID"]

    def convert_pdf_to_docx(self, access_token, asset_id):
        url = "https://pdf-services.adobe.io/operation/exportpdf"
        headers = {
            "x-api-key": self.credentials["client_credentials"]["client_id"],
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        response = requests.post(url, json={
            "assetID": asset_id,
            "targetFormat": "docx",
            "ocrLang": "en-US"
        }, headers=headers, timeout=30)
        response.raise_for_status()
        
        job_id = response.headers.get("Location", "").split("/")[-2]
        if not job_id:
            raise ValueError("Job ID not found in response")
        return job_id

    def poll_and_download_result(self, access_token, job_id, output_path):
        url = f"https://pdf-services.adobe.io/operation/exportpdf/{job_id}/status"
        headers = {
            "x-api-key": self.credentials["client_credentials"]["client_id"],
            "Authorization": f"Bearer {access_token}"
        }
        retries = 0
        download_uri = None
        
        while retries < MAX_RETRIES and self._is_running:
            try:
                response = requests.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                status_data = response.json()
                
                if status_data.get("status") == "done":
                    download_uri = status_data.get("downloadUri") or status_data.get("asset", {}).get("downloadUri")
                    if not download_uri:
                        raise ValueError("Download URI not found")
                    break
                
                time.sleep(POLL_INTERVAL)
                retries += 1
            except Exception as e:
                if retries >= MAX_RETRIES:
                    raise e
                retries += 1
                time.sleep(POLL_INTERVAL)
        
        if not download_uri:
            raise ValueError("Conversion completed but download URI is missing")
        
        response = requests.get(download_uri, stream=True, timeout=UPLOAD_TIMEOUT)
        response.raise_for_status()
        with open(output_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if not self._is_running:
                    break
                f.write(chunk)    
    
    def stop(self):
        self._is_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.workers = []
        self.threads = []
        self.credentials_path = None
        self.settings = QSettings("PDFConverter", "PDFtoDOCX")  # For Remember Me feature
        self.initUI()
        self.load_credentials()
        self.setStyleSheet(STYLESHEET)
        self.setWindowIcon(QIcon(self.resource_path("favicon.ico")))  # Set favicon
        
    def resource_path(self, relative_path):
        """Get the absolute path to a resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)        

    def initUI(self):
        self.setWindowTitle("PDF to DOCX Converter Pro")
        self.setGeometry(100, 100, 1000, 600)

        # Central widget setup
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # Left panel - File Selection
        left_panel = QVBoxLayout()
        left_panel.setSpacing(15)
        
        # Credentials selection
        credentials_layout = QHBoxLayout()
        self.btn_choose_key = QPushButton("🔑 Choose Key")
        self.btn_choose_key.clicked.connect(self.choose_credentials)
        credentials_layout.addWidget(self.btn_choose_key)
        
        # File selection controls
        control_layout = QHBoxLayout()
        self.btn_select = QPushButton("📁 Select PDFs")
        self.btn_select.clicked.connect(self.select_files)
        self.btn_clear = QPushButton("❌ Clear All")
        self.btn_clear.clicked.connect(self.clear_all_files)
        control_layout.addWidget(self.btn_select)
        control_layout.addWidget(self.btn_clear)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.setFont(QFont("Segoe UI", 10))
        
        left_panel.addLayout(credentials_layout)
        left_panel.addLayout(control_layout)
        left_panel.addWidget(QLabel("Selected Files:"))
        left_panel.addWidget(self.file_list)
        
        # Right panel - Progress Tracking
        right_panel = QVBoxLayout()
        right_panel.setSpacing(15)
        
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["File Name", "Progress", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setFont(QFont("Segoe UI", 10))
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Action buttons
        button_layout = QHBoxLayout()
        self.btn_start = QPushButton("🚀 Start Conversion")
        self.btn_stop = QPushButton("⛔ Stop All")
        self.btn_start.clicked.connect(self.start_conversion)
        self.btn_stop.clicked.connect(self.stop_all)
        button_layout.addWidget(self.btn_start)
        button_layout.addWidget(self.btn_stop)
        
        right_panel.addWidget(QLabel("Conversion Progress:"))
        right_panel.addWidget(self.table)
        right_panel.addLayout(button_layout)

        # Add panels to main layout
        main_layout.addLayout(left_panel, 40)
        main_layout.addLayout(right_panel, 60)

        # Set custom font
        font = QFont("Segoe UI", 10)
        self.setFont(font)
        
    def choose_credentials(self):
        """Allow user to select the credentials file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose Credentials File", "", "JSON Files (*.json)"
        )
        if file_path:
            self.credentials_path = file_path
            self.settings.setValue("credentials_path", self.credentials_path)  # Save path
            self.load_credentials()
            
    def load_credentials(self):
        """Load credentials from the selected file or last saved path"""
        self.credentials_path = self.settings.value("credentials_path")  # Load saved path
        if self.credentials_path and os.path.exists(self.credentials_path):
            try:
                with open(self.credentials_path, "r") as f:
                    self.credentials = json.load(f)
                print("Credentials loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load credentials: {e}")
        else:
            QMessageBox.warning(self, "Warning", "No credentials file selected.")
            

    def select_files(self):
        """Open file dialog to select PDF files"""
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select PDF Files", 
            "", 
            "PDF Files (*.pdf)"
        )
        if files:
            self.file_list.addItems(files)
            for file in files:
                self.add_progress_row(file)                                
        
    def select_files(self):
        """Open file dialog to select PDF files"""
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select PDF Files", 
            "", 
            "PDF Files (*.pdf)"
        )
        if files:
            self.file_list.addItems(files)
            for file in files:
                self.add_progress_row(file)        

    def clear_all_files(self):
        """Clear all selected files and progress"""
        self.file_list.clear()
        self.table.setRowCount(0)
        self.stop_all()

    def load_credentials(self):
        try:
            with open("X:/adobe/pdfservices-api-credentials.json", "r") as f:
                self.credentials = json.load(f)
        except Exception as e:
            print(f"Error loading credentials: {e}")

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        if files:
            self.file_list.addItems(files)
            for file in files:
                self.add_progress_row(file)

    def add_progress_row(self, file_path):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(Path(file_path).name))
        progress = QProgressBar()
        progress.setValue(0)
        self.table.setCellWidget(row, 1, progress)
        self.table.setItem(row, 2, QTableWidgetItem("Pending"))

    def start_conversion(self):
        for i in range(self.file_list.count()):
            file_path = self.file_list.item(i).text()
            self.start_worker(file_path)

    def start_worker(self, file_path):
        worker = ConversionWorker(file_path, self.credentials)
        thread = QThread()
        worker.moveToThread(thread)

        worker.progress_updated.connect(self.update_progress)
        worker.finished.connect(self.conversion_finished)
        thread.started.connect(worker.run)
        worker.stopped.connect(thread.quit)
        worker.stopped.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        self.workers.append(worker)
        self.threads.append(thread)
        thread.start()

    def update_progress(self, file_path, progress):
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).text() == Path(file_path).name:
                progress_bar = self.table.cellWidget(row, 1)
                progress_bar.setValue(progress)
                
                # Update progress bar color based on percentage
                if progress < 33:
                    color = COLORS['danger']
                elif progress < 66:
                    color = COLORS['primary']
                else:
                    color = COLORS['success']
                
                progress_bar.setStyleSheet(f"""
                    QProgressBar::chunk {{
                        background-color: {color};
                    }}
                """)
                
    def conversion_finished(self, file_path, success, message):
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).text() == Path(file_path).name:
                status_item = self.table.item(row, 2)
                if success:
                    status_item.setText("✅ Completed")
                    status_item.setForeground(QColor(COLORS['success']))
                else:
                    status_item.setText(f"❌ Failed: {message}")
                    status_item.setForeground(QColor(COLORS['danger']))

    def stop_all(self):
        for worker in self.workers:
            worker.stop()
        self.workers.clear()
        self.threads.clear()
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())