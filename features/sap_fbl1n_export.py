# features/sap_fbl1n_export.py
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QFormLayout, QLineEdit,
    QPushButton, QTextEdit, QHBoxLayout, QDateEdit, QGroupBox,
    QComboBox
)
from PyQt5.QtCore import QDate, Qt
import os
import json
import time
import pythoncom
import win32com.client

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')

class SapFBL1NExportWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # 标题
        title = QLabel("📊 海关 供应商明细账导出")
        title.setObjectName("titleLabel")
        layout.addWidget(title)
        
        # 海关配置组
        fbl1n_group = QGroupBox("海关查询参数")
        fbl1n_layout = QFormLayout(fbl1n_group)
        
        # 供应商代码
        self.vendor_input = QLineEdit()
        self.vendor_input.setPlaceholderText("请输入供应商代码")
        fbl1n_layout.addRow("供应商代码:", self.vendor_input)
        
        # 公司代码
        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText("请输入公司代码")
        fbl1n_layout.addRow("公司代码:", self.company_input)
        
        # 日期选择
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        fbl1n_layout.addRow("截止日期:", self.date_edit)
        
        # 导出格式
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Excel (.xlsx)", "CSV (.csv)", "PDF (.pdf)"])
        fbl1n_layout.addRow("导出格式:", self.format_combo)
        
        layout.addWidget(fbl1n_group)
        
        # 导出选项组
        export_group = QGroupBox("导出选项")
        export_layout = QFormLayout(export_group)
        
        # 导出路径
        self.export_path = QLineEdit()
        self.export_path.setPlaceholderText("请选择导出路径")
        export_layout.addRow("导出路径:", self.export_path)
        
        # 文件名前缀
        self.filename_prefix = QLineEdit()
        self.filename_prefix.setPlaceholderText("请输入文件名前缀")
        export_layout.addRow("文件名前缀:", self.filename_prefix)
        
        layout.addWidget(export_group)
        
        # 运行按钮
        self.run_button = QPushButton("🚀 运行 海关 导出")
        self.run_button.clicked.connect(self.run_fbl1n_export)
        layout.addWidget(self.run_button)
        
        # 日志显示
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)
        
        # 加载配置
        self.load_config()
        
    def load_config(self):
        """加载保存的海关配置"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "fbl1n" in data:
                        fbl1n_data = data["fbl1n"]
                        vendor = fbl1n_data.get("vendor", "")
                        if vendor:
                            self.vendor_input.setText(vendor)
                        company = fbl1n_data.get("company", "")
                        if company:
                            self.company_input.setText(company)
                        export_path = fbl1n_data.get("export_path", "")
                        if export_path:
                            self.export_path.setText(export_path)
                        prefix = fbl1n_data.get("filename_prefix", "")
                        if prefix:
                            self.filename_prefix.setText(prefix)
            except Exception as e:
                self.log_message(f"加载配置时出错: {str(e)}")
    
    def save_config(self):
        """保存海关配置"""
        try:
            data = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            
            # 更新或创建FBL1N配置
            data["fbl1n"] = {
                "vendor": self.vendor_input.text(),
                "company": self.company_input.text(),
                "export_path": self.export_path.text(),
                "filename_prefix": self.filename_prefix.text()
            }
            
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
            self.log_message("✅ 配置已保存")
        except Exception as e:
            self.log_message(f"❌ 保存配置时出错: {str(e)}")
    
    def log_message(self, message):
        """添加日志消息"""
        self.log.append(message)
        
    def run_fbl1n_export(self):
        """运行海关导出"""
        self.log.clear()
        self.log_message("⏳ 开始执行海关导出...")
        
        # 保存当前配置
        self.save_config()
        
        # 获取参数
        vendor = self.vendor_input.text()
        company = self.company_input.text()
        selected_date = self.date_edit.date().toString("yyyyMMdd")
        formatted_date = self.date_edit.date().toString("yyyy-MM-dd")
        export_format = self.format_combo.currentText().split("(")[0].strip()
        
        self.log_message(f"供应商代码: {vendor}")
        self.log_message(f"公司代码: {company}")
        self.log_message(f"截止日期: {formatted_date}")
        self.log_message(f"导出格式: {export_format}")
        
        # 初始化COM
        pythoncom.CoInitialize()
        
        try:
            # 获取SAP GUI对象
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
                self.log_message("✅ 已连接到SAP GUI")
            except Exception:
                self.log_message("❌ SAP GUI 未找到，请确保 SAP GUI 已打开并登录")
                return
            
            # 获取应用程序和会话
            try:
                application = SapGuiAuto.GetScriptingEngine
                connection = application.Children(0)
                session = connection.Children(0)
                self.log_message("✅ 已获取SAP会话")
            except Exception:
                self.log_message("❌ 获取 SAP Session 失败，请检查 SAP 是否已登录")
                return
            
            # 执行海关导出
            try:
                # 最大化窗口
                session.findById("wnd[0]").maximize()
                
                # 输入事务代码FBL1N
                self.log_message("⏳ 正在打开海关事务...")
                session.findById("wnd[0]/tbar[0]/okcd").text = "fbl1n"
                session.findById("wnd[0]").sendVKey(0)
                
                # 设置供应商明细账选项
                self.log_message("⏳ 正在设置查询参数...")
                session.findById("wnd[0]/usr/chkX_SHBV").selected = True
                session.findById("wnd[0]/usr/chkX_MERK").selected = True
                
                # 输入供应商代码
                session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = vendor
                
                # 输入公司代码
                session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = company
                
                # 设置日期
                input_date = self.date_edit.date().toString("dd.MM.yyyy")
                session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = input_date
                
                # 执行查询
                self.log_message("⏳ 正在执行查询...")
                session.findById("wnd[0]").sendVKey(8)  # F8 执行
                
                # 导出到Excel - 使用与OpenAP相同的简化流程
                self.log_message("⏳ 正在导出数据...")
                try:
                    # 记录当前可用的窗口和控件
                    self.log_message("  - 调试: 记录当前窗口状态")
                    try:
                        windows = []
                        for i in range(session.Children.Count):
                            windows.append(f"wnd[{i}]")
                        self.log_message(f"  - 可用窗口: {', '.join(windows)}")
                    except Exception as e:
                        self.log_message(f"  - 无法获取窗口列表: {str(e)}")
                    
                    self.log_message("  - 选择菜单: 列表 > 导出 > 本地文件")
                    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()  # 列表 > 导出 > 本地文件
                except Exception as e:
                    self.log_message(f"❌ 选择导出菜单失败: {str(e)}")
                    raise
                
                # 直接按确定按钮，跳过选择格式步骤
                try:
                    self.log_message("  - 确认导出对话框")
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception as e:
                    self.log_message(f"❌ 确认导出对话框失败: {str(e)}")
                    # 尝试记录当前对话框的控件
                    try:
                        self.log_message("  - 调试: 尝试识别当前对话框控件")
                        dialog = session.findById("wnd[1]")
                        for i in range(dialog.Children.Count):
                            child = dialog.Children(i)
                            self.log_message(f"  - 控件 {i}: {child.Id}")
                    except Exception as debug_e:
                        self.log_message(f"  - 无法获取对话框控件: {str(debug_e)}")
                    raise
                
                # 设置保存路径和文件名 - 与OpenAP保持一致
                try:
                    # 使用用户指定的路径和文件名前缀
                    desktop_path = self.export_path.text()
                    user_prefix = self.filename_prefix.text()
                    filename = f"{user_prefix}_{selected_date}.xlsx"
                    full_path = os.path.join(desktop_path, filename)
                    
                    self.log_message("  - 设置保存路径和文件名")
                    self.log_message(f"  - 路径: {desktop_path}")
                    self.log_message(f"  - 文件名: {filename}")
                    
                    # 尝试获取当前对话框的控件ID
                    try:
                        dialog = session.findById("wnd[1]")
                        for i in range(dialog.Children.Count):
                            try:
                                child = dialog.Children(i)
                                self.log_message(f"  - 控件 {i}: {child.Id}")
                            except:
                                pass
                    except Exception as debug_e:
                        self.log_message(f"  - 调试: 无法列出控件: {str(debug_e)}")
                    
                    # 使用与OpenAP相同的控件ID
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = desktop_path
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                    
                    self.log_message("  - 确认保存")
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    
                    # 检查文件是否成功创建
                    time.sleep(2)  # 等待文件保存完成
                    if os.path.exists(full_path):
                        self.log_message(f"✅ 文件已成功保存到: {full_path}")
                    else:
                        self.log_message(f"⚠️ 文件可能未成功保存，请检查路径: {full_path}")
                        # 检查桌面上是否有其他类似名称的文件
                        try:
                            import glob
                            user_prefix = self.filename_prefix.text()
                            similar_files = glob.glob(os.path.join(desktop_path, f"{user_prefix}_*.xlsx"))
                            if similar_files:
                                self.log_message(f"  - 找到类似文件: {', '.join([os.path.basename(f) for f in similar_files])}")
                        except Exception as glob_e:
                            self.log_message(f"  - 无法检查类似文件: {str(glob_e)}")
                except Exception as e:
                    self.log_message(f"❌ 保存文件时出错: {str(e)}")
                    # 尝试使用替代方法
                    try:
                        self.log_message("  - 尝试替代保存方法...")
                        # 尝试使用键盘快捷键或其他方法保存
                        session.findById("wnd[1]").sendVKey(11)  # F11通常是保存
                        self.log_message("  - 已尝试使用F11键保存")
                    except Exception as alt_e:
                        self.log_message(f"  - 替代保存方法失败: {str(alt_e)}")
                    raise
                
            except Exception as e:
                self.log_message(f"❌ 执行过程中出现错误: {str(e)}")
        finally:
            pythoncom.CoUninitialize()