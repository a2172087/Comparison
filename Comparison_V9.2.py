import os
import sys
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QFileDialog, QProgressBar, QDesktopWidget, QMessageBox, QCheckBox
import qtmodern.styles
import qtmodern.windows
import shutil
from datetime import datetime
import time
import socket
import re
import py7zr
import tempfile
import datetime

def get_application_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))

def set_app_style(app):
    font = QFont("微軟正黑體", 9)
    font.setBold(True)
    app.setFont(font)
    qtmodern.styles.dark(app)

class ComparisonThread(QThread):
    progress_updated = pyqtSignal(int, str)
    comparison_completed = pyqtSignal()

    def __init__(self, folder1_path, folder2_path, move_overkill_files):
        super().__init__()
        self.folder1_path = folder1_path
        self.folder2_path = folder2_path
        self.move_overkill_files = move_overkill_files

    def run(self):
        start_time = time.time()

        folder1_dirs = self.get_directories(self.folder1_path)
        folder2_dirs = self.get_directories(self.folder2_path)

        total_dirs = len(folder1_dirs)
        processed_dirs = 0

        for dir1 in folder1_dirs:
            dir1_name = os.path.basename(dir1)

            for dir2 in folder2_dirs:
                dir2_name = os.path.basename(dir2)

                if dir1_name == dir2_name:
                    self.process_directory(dir1, dir2)
                    break

            processed_dirs += 1
            progress = int((processed_dirs / total_dirs) * 100)
            elapsed_time = time.time() - start_time
            remaining_time = (elapsed_time / processed_dirs) * (total_dirs - processed_dirs)
            progress_text = f"比對進度: {processed_dirs}/{total_dirs},預計剩餘時間: {int(remaining_time)}秒"
            self.progress_updated.emit(progress, progress_text)

            QThread.msleep(100)

        self.comparison_completed.emit()

    def get_directories(self, folder_path):
        directories = []

        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isdir(item_path):
                directories.append(item_path)

        return directories

    def process_directory(self, dir1, dir2):
        for file_name in os.listdir(dir1):
            file1_path = os.path.join(dir1, file_name)

            if os.path.isfile(file1_path):
                file2_path = self.find_file_in_directory(dir2, file_name)

                if file2_path:
                    category_dir = os.path.dirname(file2_path)
                    category_name = os.path.basename(category_dir)

                    if "over" in category_name.lower() and self.move_overkill_files:
                        continue

                    target_dir = os.path.join(dir1, category_name)
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)

                    target_file_path = os.path.join(target_dir, file_name)
                    shutil.move(file1_path, target_file_path)

    def find_file_in_directory(self, directory, file_name):
        for root, dirs, files in os.walk(directory):
            if file_name in files:
                return os.path.join(root, file_name)

        return None

class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Comparison')
        self.user_id = self.get_user_id()
        self.log_folder = r"M:\Q200\02 QE Personal Data\Racky\Log History\Comparison"
        self.log_path = os.path.join(self.log_folder, f"{self.user_id}.txt")
        self.write_log()
        self.setup_ui()
        self.set_window_size(800, 300)
        self.center_window()

    def setup_ui(self):
        layout = QVBoxLayout()

        # 資料夾選擇按鈕
        self.folder1_btn = QPushButton('資料夾')
        self.folder1_btn.clicked.connect(self.select_folder1)
        layout.addWidget(self.folder1_btn)

        # 顯示選擇的資料夾路徑
        self.folder1_label = QLabel('未選擇資料夾')
        layout.addWidget(self.folder1_label)

        # 比對的資料夾選擇按鈕
        self.folder2_btn = QPushButton('比對的資料夾')
        self.folder2_btn.clicked.connect(self.select_folder2)
        layout.addWidget(self.folder2_btn)

        # 顯示選擇的比對資料夾路徑
        self.folder2_label = QLabel('未選擇資料夾')
        layout.addWidget(self.folder2_label)

        # Over Kill照片需重新Review勾選框
        self.overkill_checkbox = QCheckBox('Over Kill照片需重新Review')
        self.overkill_checkbox.setChecked(True)
        layout.addWidget(self.overkill_checkbox)

        # 執行按鈕
        self.execute_btn = QPushButton('執行')
        self.execute_btn.clicked.connect(self.execute_comparison)
        layout.addWidget(self.execute_btn)

        # 進度條
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # 進度文字
        self.progress_label = QLabel()
        layout.addWidget(self.progress_label)

        self.setLayout(layout)

    def set_window_size(self, width, height):
        self.setFixedSize(width, height)

    def center_window(self):
        screen_geometry = QDesktopWidget().availableGeometry()
        window_geometry = self.geometry()
        center_point = screen_geometry.center()
        window_geometry.moveCenter(center_point)
        self.move(window_geometry.topLeft())

    def select_folder1(self):
        folder_path = QFileDialog.getExistingDirectory(self, '選擇資料夾')
        if folder_path:
            self.folder1_label.setText(folder_path)

    def select_folder2(self):
        folder_path = QFileDialog.getExistingDirectory(self, '選擇比對的資料夾')
        if folder_path:
            self.folder2_label.setText(folder_path)

    def execute_comparison(self):
        folder1_path = self.folder1_label.text()
        folder2_path = self.folder2_label.text()
        move_overkill_files = self.overkill_checkbox.isChecked()

        if folder1_path != '未選擇資料夾' and folder2_path != '未選擇資料夾':
            self.execute_btn.setEnabled(False)
            self.execute_btn.setText('執行中...')
            self.comparison_thread = ComparisonThread(folder1_path, folder2_path, move_overkill_files)
            self.comparison_thread.progress_updated.connect(self.update_progress)
            self.comparison_thread.comparison_completed.connect(self.show_completion_message)
            self.comparison_thread.start()

    def update_progress(self, progress, progress_text):
        self.progress_bar.setValue(progress)
        self.progress_label.setText(progress_text)

    def show_completion_message(self):
        self.execute_btn.setEnabled(True)
        self.execute_btn.setText('執行')
        self.progress_bar.setValue(100)
        self.progress_label.setText('比對完成')
        QMessageBox.information(self, 'Comparison', '比對完成')

    def get_user_id(self):
        device_name = os.environ['COMPUTERNAME']
        user_id = device_name.rstrip("W10").rstrip("W11")
        return user_id

    def write_log(self):
        try:
            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            username = match.group(1) if match else 'Unknown'

            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_folder = r'M:\QA_Program_Raw_Data\Log History'
            archive_path = os.path.join(log_folder, 'Comparison.7z')
            log_filename = f'{username}.txt'
            new_log_message = f"{current_datetime} {username} Open\n"
            os.makedirs(log_folder, exist_ok=True)

            if not os.path.exists(archive_path):
                with py7zr.SevenZipFile(archive_path, mode='w', password='@Joe11111111') as archive:
                    archive.writestr(new_log_message, f'Comparison/{log_filename}')
            else:
                log_content = ""
                files_to_keep = []

                with py7zr.SevenZipFile(archive_path, mode='r', password='@Joe11111111') as archive:
                    for filename, bio in archive.read().items():
                        if filename == f'Comparison/{log_filename}':
                            log_content = bio.read().decode('utf-8')
                        else:
                            files_to_keep.append((filename, bio.read()))

                if new_log_message not in log_content:
                    log_content += new_log_message

                with tempfile.NamedTemporaryFile(delete=False, suffix='.7z') as temp_file:
                    temp_archive_path = temp_file.name

                with py7zr.SevenZipFile(temp_archive_path, mode='w', password='@Joe11111111') as archive:
                    archive.writestr(log_content.encode('utf-8'), f'Comparison/{log_filename}')
                    for filename, content in files_to_keep:
                        archive.writestr(content, filename)

                shutil.move(temp_archive_path, archive_path)

        except Exception as e:
            print(f"寫入log時發生錯誤: {e}")

    def check_latest_version(self):
        try:
            app_folder = r"M:\QA_Program_Raw_Data\Apps"
            exe_files = [f for f in os.listdir(app_folder) if f.startswith("Comparison_V") and f.endswith(".exe")]

            if not exe_files:
                QMessageBox.warnin(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                sys.exit(1)

            # 修改版本號提取邏輯，只取主版本號
            latest_version = max(int(re.search(r'_V(\d+)', f).group(1)) for f in exe_files)

            # 修改當前版本號提取邏輯，只取主版本號
            current_version_match = re.search(r'_V(\d+)', os.path.basename(sys.executable))
            if current_version_match:
                current_version = int(current_version_match.group(1))
            else:
                current_version = 0

            if current_version < latest_version:
                QMessageBox.information(self, '請更新至最新版本', '請更新至最新版本')
                os.startfile(app_folder)  # 開啟指定的資料夾
                sys.exit(0)

            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            if match:
                username = match.group(1)
                if username == "A000000":
                    QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                    sys.exit(1)
            else:
                QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                sys.exit(1)

        except FileNotFoundError:
            QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
            sys.exit(1)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    application_path = get_application_path()
    icon_path = os.path.join(application_path, 'format.ico')
    app.setWindowIcon(QIcon(icon_path))
    set_app_style(app)

    window = ImageClassifier()
    window.check_latest_version()
    win = qtmodern.windows.ModernWindow(window)
    win.show()

    sys.exit(app.exec_())