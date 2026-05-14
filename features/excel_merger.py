import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QFileDialog, QTextEdit, QProgressBar, QMessageBox, 
    QCheckBox, QComboBox, QGroupBox, QRadioButton
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


class ExcelMergeThread(QThread):
    progress_updated = pyqtSignal(int)
    merge_finished = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, folder_path, merge_type, skip_header, sheet_name=None):
        super().__init__()
        self.folder_path = folder_path
        self.merge_type = merge_type
        self.skip_header = skip_header
        self.sheet_name = sheet_name

    def run(self):
        try:
            excel_files = [f for f in os.listdir(self.folder_path) 
                          if f.endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                self.error_occurred.emit("未找到Excel文件")
                return

            all_data = []
            total_files = len(excel_files)

            for i, file in enumerate(excel_files):
                file_path = os.path.join(self.folder_path, file)
                try:
                    # 添加dtype=str参数以保留大数字的精度
                    if self.merge_type == "按工作表合并":
                        if self.sheet_name:
                            df = pd.read_excel(file_path, sheet_name=self.sheet_name, dtype=str)
                        else:
                            df = pd.read_excel(file_path, sheet_name=0, dtype=str)
                    else:
                        xls = pd.ExcelFile(file_path)
                        dfs = []
                        for sheet in xls.sheet_names:
                            sheet_df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
                            sheet_df['来源文件'] = file
                            sheet_df['来源工作表'] = sheet
                            dfs.append(sheet_df)
                        df = pd.concat(dfs, ignore_index=True)

                    if self.merge_type == "按工作表合并":
                        df['来源文件'] = file

                    all_data.append(df)
                    
                except Exception as e:
                    self.error_occurred.emit(f"处理文件 {file} 时出错: {str(e)}")
                    continue

                self.progress_updated.emit(int((i + 1) / total_files * 100))

            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                output_path = os.path.join(self.folder_path, "合并结果.xlsx")
                # 使用float_format参数确保大数字不会被转换为科学计数法
                merged_df.to_excel(output_path, index=False, float_format='%.0f')
                self.merge_finished.emit(f"合并完成！共处理 {len(excel_files)} 个文件，结果保存在: {output_path}")
            else:
                self.error_occurred.emit("没有成功处理任何文件")

        except Exception as e:
            self.error_occurred.emit(f"合并过程中出错: {str(e)}")


class ExcelMergerWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel批量合并工具")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 文件夹选择区域
        folder_group = QGroupBox("选择文件夹")
        folder_layout = QHBoxLayout()
        
        self.folder_label = QLabel("未选择文件夹")
        self.folder_label.setStyleSheet("color: #CCCCCC;")
        
        self.select_btn = QPushButton("选择文件夹")
        self.select_btn.clicked.connect(self.select_folder)
        
        folder_layout.addWidget(self.folder_label, 1)
        folder_layout.addWidget(self.select_btn)
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)

        # 合并选项区域
        options_group = QGroupBox("合并选项")
        options_layout = QVBoxLayout()
        
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("合并方式:"))
        
        self.merge_type_combo = QComboBox()
        self.merge_type_combo.addItems(["按工作表合并", "合并所有工作表"])
        self.merge_type_combo.currentTextChanged.connect(self.on_merge_type_changed)
        type_layout.addWidget(self.merge_type_combo)
        options_layout.addLayout(type_layout)

        # 工作表选择
        self.sheet_group = QGroupBox("工作表选择")
        sheet_layout = QVBoxLayout()
        
        self.first_sheet_radio = QRadioButton("合并第一个工作表")
        self.first_sheet_radio.setChecked(True)
        
        self.specific_sheet_radio = QRadioButton("指定工作表:")
        self.sheet_name_combo = QComboBox()
        self.sheet_name_combo.setEnabled(False)
        
        self.specific_sheet_radio.toggled.connect(
            lambda checked: self.sheet_name_combo.setEnabled(checked))
        
        sheet_layout.addWidget(self.first_sheet_radio)
        sheet_layout.addWidget(self.specific_sheet_radio)
        sheet_layout.addWidget(self.sheet_name_combo)
        self.sheet_group.setLayout(sheet_layout)
        options_layout.addWidget(self.sheet_group)

        self.skip_header_check = QCheckBox("跳过表头（除第一个文件外）")
        self.skip_header_check.setChecked(True)
        options_layout.addWidget(self.skip_header_check)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 操作按钮
        button_layout = QHBoxLayout()
        
        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.clicked.connect(self.start_merge)
        self.merge_btn.setEnabled(False)
        
        self.clear_btn = QPushButton("清空日志")
        self.clear_btn.clicked.connect(self.clear_log)
        
        button_layout.addWidget(self.merge_btn)
        button_layout.addWidget(self.clear_btn)
        layout.addLayout(button_layout)

        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #CBD5E0;
                border-radius: 12px;
                margin-top: 10px;
                padding-top: 10px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 20px;
                color: #1A202C;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 20px;
                padding: 0 10px 0 10px;
            }
            QPushButton {
                background-color: #4A5568;
                border: none;
                padding: 15px 25px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #FFFFFF;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2D3748;
            }
            QPushButton:disabled {
                background-color: #A0AEC0;
                color: #F7FAFC;
            }
            QTextEdit, QComboBox {
                background-color: #FFFFFF;
                border: 2px solid #CBD5E0;
                border-radius: 8px;
                padding: 14px;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #2D3748;
            }
            QTextEdit:focus, QComboBox:focus {
                border: 2px solid #4A5568;
            }
            QLabel {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
                color: #4A5568;
                font-weight: 500;
            }
            QCheckBox {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
                color: #4A5568;
                spacing: 12px;
                font-weight: 500;
            }
            QCheckBox::indicator {
                width: 22px;
                height: 22px;
                border: 2px solid #4A5568;
                border-radius: 4px;
                background-color: #FFFFFF;
            }
            QCheckBox::indicator:checked {
                background-color: #4A5568;
                border: 2px solid #4A5568;
            }
            QCheckBox::indicator:checked:hover {
                background-color: #2D3748;
            }
            QRadioButton {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
                color: #4A5568;
                spacing: 8px;
                font-weight: 500;
            }
            QRadioButton::indicator {
                width: 20px;
                height: 20px;
                border: 2px solid #4A5568;
                border-radius: 10px;
                background-color: #FFFFFF;
            }
            QRadioButton::indicator:checked {
                background-color: #4A5568;
                border: 2px solid #4A5568;
            }
            QProgressBar {
                background-color: #CBD5E0;
                border-radius: 8px;
                height: 8px;
                text-align: center;
                color: transparent;
            }
            QProgressBar::chunk {
                background-color: #4A5568;
                border-radius: 8px;
            }
        """)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含Excel文件的文件夹")
        if folder:
            self.folder_label.setText(folder)
            self.merge_btn.setEnabled(True)
            self.scan_excel_files(folder)

    def scan_excel_files(self, folder_path):
        excel_files = [f for f in os.listdir(folder_path) 
                      if f.endswith(('.xlsx', '.xls'))]
        
        if excel_files:
            try:
                first_file = os.path.join(folder_path, excel_files[0])
                xls = pd.ExcelFile(first_file)
                self.sheet_name_combo.clear()
                self.sheet_name_combo.addItems(xls.sheet_names)
                self.log(f"扫描到 {len(excel_files)} 个Excel文件")
            except Exception as e:
                self.log(f"扫描文件时出错: {str(e)}")

    def on_merge_type_changed(self, text):
        self.sheet_group.setVisible(text == "按工作表合并")

    def start_merge(self):
        folder_path = self.folder_label.text()
        if folder_path == "未选择文件夹":
            QMessageBox.warning(self, "警告", "请先选择文件夹")
            return

        merge_type = self.merge_type_combo.currentText()
        skip_header = self.skip_header_check.isChecked()
        
        sheet_name = None
        if merge_type == "按工作表合并":
            if self.specific_sheet_radio.isChecked():
                sheet_name = self.sheet_name_combo.currentText()

        self.merge_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.merge_thread = ExcelMergeThread(folder_path, merge_type, skip_header, sheet_name)
        self.merge_thread.progress_updated.connect(self.progress_bar.setValue)
        self.merge_thread.merge_finished.connect(self.on_merge_finished)
        self.merge_thread.error_occurred.connect(self.on_error)
        self.merge_thread.start()

    def on_merge_finished(self, message):
        self.progress_bar.setVisible(False)
        self.merge_btn.setEnabled(True)
        self.log(message)
        QMessageBox.information(self, "成功", message)

    def on_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.merge_btn.setEnabled(True)
        self.log(f"错误: {error_message}")
        QMessageBox.critical(self, "错误", error_message)

    def clear_log(self):
        self.log_text.clear()

    def log(self, message):
        import pandas as pd
        self.log_text.append(f"[{pd.Timestamp.now().strftime('%H:%M:%S')}] {message}")