# features/sap_integrated.py
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QFormLayout, QLineEdit,
    QPushButton, QTextEdit, QHBoxLayout, QFileDialog, 
    QDateEdit, QSplitter, QGroupBox, QTabWidget, QSpinBox,
    QProgressBar, QMessageBox, QMenu, QAction, QApplication,
    QTableWidget, QTableWidgetItem, QHeaderView
)
from features.sap_fbl1n_export import SapFBL1NExportWidget
from PyQt5.QtCore import QDate, Qt, QThread, pyqtSignal
from datetime import datetime
import os
import json
import subprocess
import time
import pythoncom
import win32com.client
import openpyxl

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
DEFAULT_SAP_PATH = ""


class F03CustomWorker(QThread):
    progress_update = pyqtSignal(int, int, str)
    log_message = pyqtSignal(str)
    finished_signal = pyqtSignal(int, float)
    error_signal = pyqtSignal(str)
    
    def __init__(self, excel_path, start_row, end_row):
        super().__init__()
        self.excel_path = excel_path
        self.start_row = start_row
        self.end_row = end_row
        self.is_running = True
        
    def run(self):
        try:
            pythoncom.CoInitialize()
            processed_count = 0
            start_time = time.time()
            
            # 连接到SAP
            self.log_message.emit("⏳ 正在连接到SAP...")
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
                application = SapGuiAuto.GetScriptingEngine
                connection = application.Children(0)
                session = connection.Children(0)
                self.log_message.emit("✅ 已成功连接到SAP")
            except Exception as e:
                self.error_signal.emit(f"SAP连接失败: {str(e)}")
                return
            
            # 打开Excel文件
            self.log_message.emit(f"⏳ 正在打开Excel文件: {os.path.basename(self.excel_path)}")
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                workbook = excel.Workbooks.Open(self.excel_path)
                worksheet = workbook.Sheets(1)  # 使用第一个工作表
                self.log_message.emit("✅ Excel文件已打开")
            except Exception as e:
                self.error_signal.emit(f"Excel文件打开失败: {str(e)}")
                return
            
            # 获取数据范围
            if self.end_row == 0:
                # 自动检测最后一行
                row = self.start_row
                while worksheet.Cells(row, "A").Value is not None and str(worksheet.Cells(row, "A").Value).strip() != "":
                    row += 1
                self.end_row = row - 1
            
            total_rows = self.end_row - self.start_row + 1
            self.log_message.emit(f"📊 共找到 {total_rows} 条记录需要处理")
            
            # 先将所有Excel数据读取为字典列表
            self.log_message.emit("📖 正在读取Excel数据...")
            data_list = []
            
            for i in range(self.start_row, self.end_row + 1):
                if not self.is_running:
                    break
                    
                # 获取Excel数据
                amount = worksheet.Cells(i, "A").Value  # A列：金额
                date_value = worksheet.Cells(i, "B").Value  # B列：日期
                account = worksheet.Cells(i, "C").Value  # C列：账户
                document_number = worksheet.Cells(i, "D").Value  # D列：凭证号
                fixed_account = worksheet.Cells(i, "E").Value  # E列：固定科目号（从E列读取）
                text = worksheet.Cells(i, "F").Value  # F列：文本
                
                # 检查必要数据
                if not amount or not document_number or not fixed_account:
                    self.log_message.emit(f"第 {i} 行数据不完整，跳过处理")
                    continue
                
                # 创建数据字典
                data_dict = {
                    "row": i,
                    "amount": amount,
                    "date_value": date_value,
                    "account": account,
                    "document_number": document_number,
                    "fixed_account": fixed_account,
                    "text": text
                }
                
                data_list.append(data_dict)
            
            self.log_message.emit(f"📚 成功读取 {len(data_list)} 条有效数据")
            
            # 关闭Excel文件
            workbook.Close(False)
            excel.Quit()
            
            # 逐条处理数据字典
            for idx, data in enumerate(data_list):
                if not self.is_running:
                    break
                    
                try:
                    # 更新进度
                    self.progress_update.emit(idx + 1, len(data_list), f"正在处理凭证: {data['document_number']}")
                    
                    # 处理F-03事务
                    success, message = self.process_f03_transaction(
                        session, 
                        data["amount"], 
                        data["date_value"], 
                        data["account"], 
                        data["document_number"], 
                        data["fixed_account"], 
                        data["text"]
                    )
                    
                    if success:
                        self.log_message.emit(f"✅ 第 {data['row']} 行处理成功: {message}")
                    else:
                        self.log_message.emit(f"❌ 第 {data['row']} 行处理失败: {message}")
                        break
                    
                    processed_count += 1
                    
                except Exception as e:
                    self.log_message.emit(f"❌ 处理第 {data['row']} 行时出错: {str(e)}")
                    break
            
            elapsed_time = time.time() - start_time
            self.finished_signal.emit(processed_count, elapsed_time)
            
        except Exception as e:
            self.error_signal.emit(f"执行过程中出错: {str(e)}")
    
    def process_f03_transaction(self, session, amount, date_value, account, document_number, fixed_account, text):
        """处理单个F-03事务，完全按照标准SAP脚本操作流程"""
        try:
            # 进入F-03事务
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "f-03"
            session.findById("wnd[0]").sendVKey(0)
            
            # 选择清账选项
            session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").select()
            
            # 输入基本信息
            session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = str(account) if account else ""
            
            # 处理日期
            date_str = ""
            if date_value:
                try:
                    if isinstance(date_value, (datetime, type(datetime.now()))):
                        date_str = date_value.strftime("%d.%m.%Y")
                    else:
                        date_str = str(date_value).strip()
                        # 尝试转换为SAP需要的日期格式
                        if "-" in date_str:
                            date_parts = date_str.split("-")
                            if len(date_parts) == 3:
                                date_str = f"{date_parts[2]}.{date_parts[1]}.{date_parts[0]}"
                except:
                    pass
            
            if date_str:
                session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = date_str
            
            # 从配置文件读取公司代码和货币
            company_code = ""
            currency = ""
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        config_data = json.load(f)
                        company_code = config_data.get("default_company_code", "")
                        currency = config_data.get("default_currency", "")
                except:
                    pass
            
            if company_code:
                session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = company_code
            if currency:
                session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = currency
            
            session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").setFocus()
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            
            # 输入凭证号
            session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").text = str(document_number)
            session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").caretPosition = len(str(document_number))
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            session.findById("wnd[0]/tbar[1]/btn[7]").press()
            
            # 添加新项目
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "27"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = str(fixed_account) if fixed_account else ""
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus()
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = len(str(fixed_account)) if fixed_account else 7
            session.findById("wnd[0]").sendVKey(0)
            
            # 输入金额
            if amount:
                try:
                    amount_str = str(float(amount))
                    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = amount_str
                    session.findById("wnd[0]/usr/txtBSEG-DMBTR").text = amount_str
                except:
                    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = "0.00"
                    session.findById("wnd[0]/usr/txtBSEG-DMBTR").text = "0.00"
            else:
                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = "0.00"
                session.findById("wnd[0]/usr/txtBSEG-DMBTR").text = "0.00"
            
            # 输入文本
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = str(text) if text else ""
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").setFocus()
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = len(str(text)) if text else 0
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            
            # 返回主页
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            
            return True, "处理成功"
            
        except Exception as e:
            return False, f"处理失败: {str(e)}"
    
    def stop(self):
        self.is_running = False


class SapIntegratedWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.f03_worker = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        
        # 创建标签页
        self.tab_widget = QTabWidget()
        
        # 创建标签页
        self.login_tab = self.create_login_tab()
        self.open_ap_tab = self.create_open_ap_tab()
        self.f03_tab = self.create_f03_tab()
        self.fbl1n_tab = SapFBL1NExportWidget()
        
        self.tab_widget.addTab(self.login_tab, "🔐 SAP登录")
        self.tab_widget.addTab(self.open_ap_tab, "📊 Open AP")
        self.tab_widget.addTab(self.f03_tab, "🎯 F-03清账")
        self.tab_widget.addTab(self.fbl1n_tab, "📋 海关导出")
        
        layout.addWidget(self.tab_widget)
        
        # 加载配置
        self.load_config()

    def create_f03_tab(self):
        """创建F-03清账标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 标题
        title = QLabel("🎯 SAP F-03 自定义清账自动化")
        title.setObjectName("titleLabel")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2D3748; margin: 20px 0;")
        layout.addWidget(title)
        
        # 说明
        description = QLabel("此功能将根据您提供的Excel文件自动执行SAP F-03清账操作。\nExcel文件格式要求：A列=金额, B列=日期, C列=账户, D列=凭证号, E列=固定科目号, F列=文本")
        description.setWordWrap(True)
        description.setStyleSheet("color: #718096; margin-bottom: 20px; padding: 10px; background-color: #F7FAFC; border-radius: 8px;")
        layout.addWidget(description)
        
        # 配置组
        config_group = QGroupBox("基本配置")
        config_layout = QFormLayout(config_group)
        
        # Excel文件选择
        excel_layout = QHBoxLayout()
        self.f03_excel_path_input = QLineEdit()
        self.f03_excel_path_input.setPlaceholderText("选择包含清账数据的Excel文件")
        self.f03_excel_path_input.textChanged.connect(self.on_f03_excel_path_changed)
        excel_btn = QPushButton("📁 选择文件")
        excel_btn.clicked.connect(self.select_f03_excel_file)
        excel_layout.addWidget(self.f03_excel_path_input)
        excel_layout.addWidget(excel_btn)
        config_layout.addRow("Excel文件:", excel_layout)
        
        layout.addWidget(config_group)
        
        # 数据范围组
        range_group = QGroupBox("处理范围")
        range_layout = QFormLayout(range_group)
        
        self.f03_start_row_input = QSpinBox()
        self.f03_start_row_input.setMinimum(2)
        self.f03_start_row_input.setMaximum(10000)
        # 设置默认起始行为2（跳过标题行）
        self.f03_start_row_input.setValue(2)
        
        self.f03_end_row_input = QSpinBox()
        self.f03_end_row_input.setMinimum(0)
        self.f03_end_row_input.setMaximum(10000)
        self.f03_end_row_input.setValue(0)
        self.f03_end_row_input.setSpecialValueText("自动检测")
        
        range_layout.addRow("起始行:", self.f03_start_row_input)
        range_layout.addRow("结束行:", self.f03_end_row_input)
        
        layout.addWidget(range_group)
        
        # 控制按钮
        btn_layout = QHBoxLayout()
        self.f03_start_btn = QPushButton("🚀 开始处理")
        self.f03_start_btn.clicked.connect(self.start_f03_processing)
        self.f03_start_btn.setStyleSheet("""
            QPushButton {
                background-color: #48BB78;
                font-size: 18px;
                padding: 12px 24px;
                border-radius: 8px;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #38A169;
            }
            QPushButton:disabled {
                background-color: #CBD5E0;
            }
        """)
        
        self.f03_stop_btn = QPushButton("⏹️ 停止处理")
        self.f03_stop_btn.clicked.connect(self.stop_f03_processing)
        self.f03_stop_btn.setEnabled(False)
        self.f03_stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #F56565;
                font-size: 18px;
                padding: 12px 24px;
                border-radius: 8px;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E53E3E;
            }
            QPushButton:disabled {
                background-color: #CBD5E0;
            }
        """)
        
        btn_layout.addWidget(self.f03_start_btn)
        btn_layout.addWidget(self.f03_stop_btn)
        layout.addLayout(btn_layout)
        
        # 进度条
        self.f03_progress_bar = QProgressBar()
        self.f03_progress_bar.setVisible(False)
        layout.addWidget(self.f03_progress_bar)
        
        # 状态标签
        self.f03_status_label = QLabel("就绪")
        self.f03_status_label.setStyleSheet("color: #718096; margin: 10px 0;")
        layout.addWidget(self.f03_status_label)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        
        self.f03_log_text = QTextEdit()
        self.f03_log_text.setReadOnly(True)
        self.f03_log_text.setMaximumHeight(200)
        log_layout.addWidget(self.f03_log_text)
        
        layout.addWidget(log_group)
        
        # Excel预览表格
        preview_group = QGroupBox("Excel数据预览")
        preview_layout = QVBoxLayout(preview_group)
        
        self.f03_preview_table = QTableWidget()
        self.f03_preview_table.setColumnCount(6)
        self.f03_preview_table.setHorizontalHeaderLabels(["金额", "日期", "账户", "凭证号", "固定科目号", "文本"])
        self.f03_preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.f03_preview_table.setMaximumHeight(150)
        preview_layout.addWidget(self.f03_preview_table)
        
        layout.addWidget(preview_group)
        
        return tab

    def create_login_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # 登录配置组
        login_group = QGroupBox("SAP登录配置")
        login_layout = QFormLayout(login_group)
        
        self.input_sap_path = QLineEdit()
        btn_browse = QPushButton("选择 SAP 路径")
        btn_browse.clicked.connect(self.select_sap_path)

        path_layout = QHBoxLayout()
        path_layout.addWidget(self.input_sap_path)
        path_layout.addWidget(btn_browse)
        login_layout.addRow("SAP路径:", path_layout)

        self.input_client = QLineEdit()
        self.input_user = QLineEdit()
        self.input_password = QLineEdit()
        self.input_password.setEchoMode(QLineEdit.Password)
        self.input_language = QLineEdit()
        self.input_connection_name = QLineEdit()
        self.input_connection_name.setPlaceholderText("请输入SAP连接名称")
        self.input_user_id = QLineEdit()
        self.input_user_id.setPlaceholderText("用于FBL1N查询的用户ID")

        login_layout.addRow("Client:", self.input_client)
        login_layout.addRow("User:", self.input_user)
        login_layout.addRow("Password:", self.input_password)
        login_layout.addRow("Language:", self.input_language)
        login_layout.addRow("连接名称:", self.input_connection_name)
        login_layout.addRow("用户ID:", self.input_user_id)

        layout.addWidget(login_group)

        # 登录按钮
        btn_login = QPushButton("🚀 登录 SAP")
        btn_login.clicked.connect(self.sap_login)
        layout.addWidget(btn_login)

        # 日志显示
        self.login_log = QTextEdit()
        self.login_log.setReadOnly(True)
        self.login_log.setMaximumHeight(150)
        layout.addWidget(self.login_log)

        return tab

    def create_open_ap_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Open AP配置组
        ap_group = QGroupBox("Open AP配置")
        ap_layout = QFormLayout(ap_group)
        
        # 日期选择
        date_label = QLabel("选择日期:")
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        
        ap_layout.addRow(date_label, self.date_edit)

        layout.addWidget(ap_group)

        # 运行按钮
        btn_run = QPushButton("运行 Open AP")
        btn_run.clicked.connect(self.run_open_ap)
        layout.addWidget(btn_run)

        # 日志显示
        self.ap_log = QTextEdit()
        self.ap_log.setReadOnly(True)
        layout.addWidget(self.ap_log)

        return tab

    def select_f03_excel_file(self):
        """选择F-03清账Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.f03_excel_path_input.setText(file_path)
            self.preview_f03_excel_data(file_path)

    def on_f03_excel_path_changed(self):
        """当Excel路径改变时预览数据"""
        file_path = self.f03_excel_path_input.text()
        if file_path and os.path.exists(file_path):
            self.preview_f03_excel_data(file_path)

    def preview_f03_excel_data(self, file_path):
        """预览Excel数据"""
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            worksheet = workbook.active
            
            # 清空表格
            self.f03_preview_table.setRowCount(0)
            
            # 显示前5行数据
            for i, row in enumerate(worksheet.iter_rows(min_row=1, max_row=6, values_only=True)):
                if i == 0:  # 跳过标题行
                    continue
                    
                row_position = self.f03_preview_table.rowCount()
                self.f03_preview_table.insertRow(row_position)
                
                for j, cell_value in enumerate(row):
                    if j < 6:  # 只显示前6列
                        item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                        self.f03_preview_table.setItem(row_position, j, item)
            
            workbook.close()
        except Exception as e:
            self.f03_log_message(f"❌ 预览Excel文件失败: {str(e)}")

    def start_f03_processing(self):
        """开始F-03处理"""
        excel_path = self.f03_excel_path_input.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "警告", "请选择Excel文件")
            return
        
        if not os.path.exists(excel_path):
            QMessageBox.warning(self, "警告", "Excel文件不存在")
            return
        
        # 获取处理范围
        start_row = self.f03_start_row_input.value()
        end_row = self.f03_end_row_input.value()
        
        # 创建并启动工作线程
        self.f03_worker = F03CustomWorker(excel_path, start_row, end_row)
        self.f03_worker.progress_update.connect(self.update_f03_progress)
        self.f03_worker.log_message.connect(self.f03_log_message)
        self.f03_worker.finished_signal.connect(self.on_f03_finished)
        self.f03_worker.error_signal.connect(self.on_f03_error)
        
        # 更新UI状态
        self.f03_start_btn.setEnabled(False)
        self.f03_stop_btn.setEnabled(True)
        self.f03_progress_bar.setVisible(True)
        self.f03_progress_bar.setValue(0)
        self.f03_status_label.setText("处理中...")
        
        # 启动线程
        self.f03_worker.start()

    def stop_f03_processing(self):
        """停止F-03处理"""
        if self.f03_worker and self.f03_worker.isRunning():
            self.f03_worker.stop()
            self.f03_log_message("⏹️ 正在停止处理...")

    def update_f03_progress(self, current, total, message):
        """更新F-03处理进度"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.f03_progress_bar.setValue(progress)
        self.f03_status_label.setText(f"处理进度: {current}/{total} - {message}")

    def f03_log_message(self, message):
        """记录F-03处理日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.f03_log_text.append(f"[{timestamp}] {message}")
        # 自动滚动到底部
        scrollbar = self.f03_log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_f03_finished(self, processed_count, elapsed_time):
        """F-03处理完成"""
        # 恢复UI状态
        self.f03_start_btn.setEnabled(True)
        self.f03_stop_btn.setEnabled(False)
        self.f03_progress_bar.setVisible(False)
        self.f03_status_label.setText(f"处理完成 - 共处理 {processed_count} 条记录，耗时 {elapsed_time:.2f} 秒")
        
        QMessageBox.information(
            self, "处理完成", 
            f"成功处理 {processed_count} 条记录\n耗时: {elapsed_time:.2f} 秒"
        )

    def on_f03_error(self, error_message):
        """F-03处理出错"""
        # 恢复UI状态
        self.f03_start_btn.setEnabled(True)
        self.f03_stop_btn.setEnabled(False)
        self.f03_progress_bar.setVisible(False)
        self.f03_status_label.setText("处理出错")
        
        self.f03_log_message(f"❌ 错误: {error_message}")
        QMessageBox.critical(self, "处理出错", error_message)

    def select_sap_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 SAP Logon 程序", "", "程序文件 (*.exe)")
        if path:
            self.input_sap_path.setText(path)
            self.save_config()

    def load_config(self):
        """加载保存的登录信息和路径"""
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.input_sap_path.setText(data.get("sap_path", DEFAULT_SAP_PATH))
                self.input_client.setText(data.get("client", ""))
                self.input_user.setText(data.get("user", ""))
                self.input_password.setText(data.get("password", ""))
                self.input_language.setText(data.get("language", ""))
                
                # 加载连接名称和用户ID
                if hasattr(self, 'input_connection_name'):
                    self.input_connection_name.setText(data.get("connection_name", ""))
                if hasattr(self, 'input_user_id'):
                    self.input_user_id.setText(data.get("user_id", ""))
                
                # 加载F-03配置
                if 'sap_f03_custom' in data:
                    settings = data['sap_f03_custom']
                    if 'last_excel_path' in settings and os.path.exists(settings['last_excel_path']):
                        self.f03_excel_path_input.setText(settings['last_excel_path'])
                        self.preview_f03_excel_data(settings['last_excel_path'])
                    if 'start_row' in settings:
                        self.f03_start_row_input.setValue(settings['start_row'])
        else:
            self.input_sap_path.setText(DEFAULT_SAP_PATH)

    def save_config(self):
        """保存登录信息和路径"""
        data = {}
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        
        # 更新SAP登录配置
        data["sap_path"] = self.input_sap_path.text()
        data["client"] = self.input_client.text()
        data["user"] = self.input_user.text()
        data["password"] = self.input_password.text()
        data["language"] = self.input_language.text()
        data["connection_name"] = self.input_connection_name.text() if hasattr(self, 'input_connection_name') else ""
        data["user_id"] = self.input_user_id.text() if hasattr(self, 'input_user_id') else ""
        
        # 保存F-03配置
        excel_path = self.f03_excel_path_input.text().strip()
        if excel_path:
            data['sap_f03_custom'] = {
                'last_excel_path': excel_path,
                'start_row': self.f03_start_row_input.value()
            }
        
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def sap_login(self):
        sap_path = self.input_sap_path.text()
        client = self.input_client.text()
        user = self.input_user.text()
        password = self.input_password.text()
        language = self.input_language.text()

        self.save_config()

        try:
            subprocess.Popen(sap_path)
            self.login_log.append("⏳ 正在启动 SAP Logon...")
            time.sleep(2)

            # 从配置文件读取连接名称
            connection_name = ""
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        config_data = json.load(f)
                        connection_name = config_data.get("connection_name", "")
                except:
                    pass
            
            if not connection_name:
                QMessageBox.warning(self, "警告", "请先在SAP登录配置中设置连接名称")
                return
            
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection(connection_name, True)
            session = connection.Children(0)

            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = client
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            session.findById("wnd[0]/usr/txtRSYST-LANGU").text = language
            session.findById("wnd[0]").sendVKey(0)

            time.sleep(2)
            if session.Children.Count > 1:
                popup = session.Children(1)
                popup.findById("usr/radMULTI_LOGON_OPT2").select()
                popup.findById("tbar[0]/btn[0]").press()

            self.login_log.append("✅ 已成功登录 SAP")
        except Exception as e:
            self.login_log.append(f"❌ 登录失败: {e}")

    def run_open_ap(self):
        selected_date = self.date_edit.date().toString("yyyyMMdd")
        formatted_date = self.date_edit.date().toString("yyyy-MM-dd")
        self.ap_log.append(f"选择的日期: {formatted_date}")
        
        pythoncom.CoInitialize()

        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            self.ap_log.append("❌ SAP GUI 未找到，请确保 SAP GUI 已打开并登录。")
            return

        try:
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
        except Exception:
            self.ap_log.append("❌ 获取 SAP Session 失败，请检查 SAP 是否已登录。")
            return

        try:
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "FBL1N"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(17)

            # 从配置文件读取用户ID，如果没有则使用默认值
            user_id = ""
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        config_data = json.load(f)
                        user_id = config_data.get("user_id", "")
                except:
                    pass
            if not user_id:
                QMessageBox.warning(self, "警告", "请先在SAP登录配置中设置用户ID")
                return
            
            session.findById("wnd[1]/usr/txtENAME-LOW").text = user_id
            session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = len(user_id)
            session.findById("wnd[1]/tbar[0]/btn[8]").press()

            try:
                grid = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
                grid.setCurrentCell(1, "TEXT")
                grid.selectedRows = 1
                grid.doubleClickCurrentCell()
            except Exception:
                self.ap_log.append("⚠️ 表格控件未找到或操作失败，跳过表格选择。")

            # 使用选择的日期
            input_date = self.date_edit.date().toString("dd.MM.yyyy")
            session.findById("wnd[0]/usr/ctxtPA_STIDA").text = input_date
            session.findById("wnd[0]/usr/ctxtPA_STIDA").caretPosition = len(input_date)

            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # 设置保存路径和文件名包含选择的日期
            # 从配置文件读取导出路径
            export_path = ""
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        config_data = json.load(f)
                        export_path = config_data.get("export_path", "")
                except:
                    pass
            
            if not export_path:
                QMessageBox.warning(self, "警告", "请先在配置中设置导出路径")
                return
            
            filename = f"AP_Report_{selected_date}.xlsx"
            full_path = os.path.join(export_path, filename)
            
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = export_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
            session.findById("wnd[1]/tbar[0]/btn[11]").press()

            self.ap_log.append(f"✅ 文件已成功保存到: {full_path}")

        except Exception as e:
            self.ap_log.append(f"❌ 执行过程中出现错误: {str(e)}")
        finally:
            pythoncom.CoUninitialize()