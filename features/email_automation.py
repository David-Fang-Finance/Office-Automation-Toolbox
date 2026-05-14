import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, 
                            QLabel, QLineEdit, QTextEdit, QPushButton, 
                            QListWidget, QListWidgetItem, QComboBox,
                            QTabWidget, QFileDialog, QMessageBox, QSplitter,
                            QGroupBox, QCheckBox, QSpinBox, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import webbrowser
import tempfile
import subprocess
import platform

class EmailDraftGenerator(QThread):
    """邮件草稿生成器 - 重构版本"""
    success_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, email_data):
        super().__init__()
        self.email_data = email_data
    
    def run(self):
        """后台执行邮件草稿创建"""
        try:
            result = self.create_draft()
            self.success_signal.emit(result)
        except Exception as e:
            self.error_signal.emit(f"邮件草稿创建失败: {str(e)}")
    
    def create_draft(self):
        """创建邮件草稿 - Windows系统静默生成"""
        try:
            # 验证必要参数
            if not self.email_data.get('to'):
                return "错误：收件人不能为空"
            if not self.email_data.get('subject'):
                return "错误：主题不能为空"
            
            # 检查系统
            system = platform.system().lower()
            if system != 'windows':
                return f"当前系统({system})不支持，此功能仅支持Windows系统"
            
            # 静默生成Outlook草稿
            return self.create_outlook_draft()
                
        except Exception as e:
            return f"草稿生成失败: {str(e)}"
    
    def create_outlook_draft(self):
        """Windows系统：静默生成Outlook草稿 - 直接保存不弹窗"""
        try:
            import win32com.client
            import pythoncom
            
            # 初始化COM库（解决-2147221008错误）
            pythoncom.CoInitialize()
            
            try:
                # 只使用COM接口静默创建，不弹窗
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                
                # 设置邮件属性
                mail.Subject = self.email_data.get('subject', '')
                mail.To = self.email_data.get('to', '')
                if self.email_data.get('cc'):
                    mail.CC = self.email_data.get('cc', '')
                
                # 处理邮件正文
                body = self.email_data.get('body', '')
                if body.strip().startswith('<') or '<html>' in body.lower():
                    mail.HTMLBody = body
                else:
                    mail.Body = body
                
                # 直接保存到草稿箱 - 不弹窗
                mail.Save()
                return f"✅ 邮件草稿已静默保存到Outlook草稿箱"
                
            finally:
                # 清理COM资源
                pythoncom.CoUninitialize()
            
        except Exception as e:
            # COM接口失败就报错，不再尝试弹窗方式
            error_msg = f"静默生成草稿失败: {str(e)}"
            print(f"Outlook COM接口错误: {error_msg}")
            return f"❌ {error_msg}"
    
    def create_macos_draft(self):
        """macOS系统：创建邮件草稿"""
        try:
            subject = self.email_data.get('subject', '')
            to = self.email_data.get('to', '')
            cc = self.email_data.get('cc', '')
            body = self.email_data.get('body', '')
            
            script = f'''
            tell application "Mail"
                set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
                tell newMessage
                    make new to recipient at end of to recipients with properties {{address:"{to}"}}
                    if "{cc}" is not "" then
                        make new cc recipient at end of cc recipients with properties {{address:"{cc}"}}
                    end if
                    save
                end tell
            end tell
            '''
            
            subprocess.run(['osascript', '-e', script], check=True)
            return "邮件草稿已保存到Mail草稿箱"
            
        except Exception:
            return self.create_mailto_draft()
    
    def create_linux_draft(self):
        """Linux系统：创建邮件草稿"""
        return self.create_mailto_draft()
    
    def create_mailto_draft(self):
        """通用mailto方式创建草稿"""
        subject = self.email_data.get('subject', '')
        to = self.email_data.get('to', '')
        body = self.html_to_text(self.email_data.get('body', ''))
        
        mailto_url = f"mailto:{to}?subject={subject}&body={body}"
        
        try:
            webbrowser.open(mailto_url)
            return "已打开默认邮件客户端创建草稿"
        except Exception:
            return "mailto方式不可用，请手动创建邮件"
    
    def html_to_text(self, html):
        """简单的HTML转文本"""
        import re
        # 移除HTML标签
        text = re.sub('<[^<]+?>', '', html)
        # 处理HTML实体
        text = text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        return text

class EmailTemplateManager:
    """邮件模板管理器"""
    
    def __init__(self):
        self.templates_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'email_templates.json')
        self.templates = self.load_templates()
    
    def load_templates(self):
        """加载邮件模板"""
        try:
            os.makedirs(os.path.dirname(self.templates_file), exist_ok=True)
            if os.path.exists(self.templates_file):
                with open(self.templates_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # 创建默认模板
                default_templates = {
                    "工作汇报": {
                        "subject": "工作汇报 - {date}",
                        "to": "",
                        "cc": "",
                        "body": "<p>尊敬的领导：</p><p>以下是本周工作汇报：</p><p>{content}</p><p>此致<br>敬礼</p>",
                        "category": "工作"
                    },
                    "会议邀请": {
                        "subject": "会议邀请 - {meeting_title}",
                        "to": "",
                        "cc": "",
                        "body": "<p>各位同事：</p><p>兹定于{date} {time}在{location}召开{meeting_title}会议。</p><p>会议议程：{agenda}</p><p>请准时参加。</p>",
                        "category": "会议"
                    },
                    "项目通知": {
                        "subject": "项目通知 - {project_name}",
                        "to": "",
                        "cc": "",
                        "body": "<p>各位项目成员：</p><p>关于{project_name}项目，有以下通知：</p><p>{notification}</p><p>请各位知悉并按此执行。</p>",
                        "category": "项目"
                    }
                }
                self.save_templates(default_templates)
                return default_templates
        except Exception as e:
            print(f"加载邮件模板失败: {str(e)}")
            return {}
    
    def save_templates(self, templates=None):
        """保存邮件模板"""
        try:
            if templates is None:
                templates = self.templates
            with open(self.templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存邮件模板失败: {str(e)}")
            return False
    
    def add_template(self, name, template_data):
        """添加邮件模板"""
        self.templates[name] = template_data
        return self.save_templates()
    
    def update_template(self, name, template_data):
        """更新邮件模板"""
        if name in self.templates:
            self.templates[name] = template_data
            return self.save_templates()
        return False
    
    def delete_template(self, name):
        """删除邮件模板"""
        if name in self.templates:
            del self.templates[name]
            return self.save_templates()
        return False
    
    def get_template(self, name):
        """获取邮件模板"""
        return self.templates.get(name, None)
    
    def get_all_templates(self):
        """获取所有邮件模板"""
        return self.templates
    
    def get_templates_by_category(self, category):
        """根据分类获取邮件模板"""
        return {name: template for name, template in self.templates.items() 
                if template.get('category') == category}

class EmailAutomationWidget(QWidget):
    """邮件自动化助手主界面"""
    
    def __init__(self):
        super().__init__()
        self.template_manager = EmailTemplateManager()
        self.init_ui()
        self.load_templates()
    
    def init_ui(self):
        """初始化界面"""
        layout = QVBoxLayout()
        
        # 创建标题
        title_label = QLabel("邮件自动化助手")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("Arial", 16, QFont.Bold)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # 创建选项卡
        self.tabs = QTabWidget()
        
        # 模板列表选项卡
        self.template_list_tab = QWidget()
        self.init_template_list_tab()
        self.tabs.addTab(self.template_list_tab, "邮件模板")
        
        # 模板编辑选项卡
        self.template_edit_tab = QWidget()
        self.init_template_edit_tab()
        self.tabs.addTab(self.template_edit_tab, "编辑模板")
        
        # 邮件生成选项卡
        self.email_generate_tab = QWidget()
        self.init_email_generate_tab()
        self.tabs.addTab(self.email_generate_tab, "生成邮件")
        
        layout.addWidget(self.tabs)
        self.setLayout(layout)
    
    def init_template_list_tab(self):
        """初始化模板列表选项卡"""
        layout = QVBoxLayout()
        
        # 分类选择
        category_layout = QHBoxLayout()
        category_layout.addWidget(QLabel("分类:"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["全部", "工作", "会议", "项目", "个人", "其他"])
        self.category_combo.currentTextChanged.connect(self.filter_templates)
        category_layout.addWidget(self.category_combo)
        category_layout.addStretch()
        layout.addLayout(category_layout)
        
        # 模板列表
        self.template_list = QListWidget()
        self.template_list.itemClicked.connect(self.on_template_selected)
        layout.addWidget(self.template_list)
        
        # 模板操作按钮
        btn_layout = QHBoxLayout()
        self.new_template_btn = QPushButton("新建模板")
        self.new_template_btn.clicked.connect(self.new_template)
        self.edit_template_btn = QPushButton("编辑模板")
        self.edit_template_btn.clicked.connect(self.edit_template)
        self.delete_template_btn = QPushButton("删除模板")
        self.delete_template_btn.clicked.connect(self.delete_template)
        self.use_template_btn = QPushButton("使用此模板")
        self.use_template_btn.clicked.connect(self.use_template)
        
        btn_layout.addWidget(self.new_template_btn)
        btn_layout.addWidget(self.edit_template_btn)
        btn_layout.addWidget(self.delete_template_btn)
        btn_layout.addWidget(self.use_template_btn)
        layout.addLayout(btn_layout)
        
        self.template_list_tab.setLayout(layout)
    
    def init_template_edit_tab(self):
        """初始化模板编辑选项卡"""
        layout = QVBoxLayout()
        
        # 模板基本信息
        info_group = QGroupBox("模板信息")
        info_layout = QGridLayout()
        
        info_layout.addWidget(QLabel("模板名称:"), 0, 0)
        self.template_name_edit = QLineEdit()
        info_layout.addWidget(self.template_name_edit, 0, 1)
        
        info_layout.addWidget(QLabel("分类:"), 1, 0)
        self.template_category_combo = QComboBox()
        self.template_category_combo.addItems(["工作", "会议", "项目", "个人", "其他"])
        info_layout.addWidget(self.template_category_combo, 1, 1)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 邮件内容
        content_group = QGroupBox("邮件内容")
        content_layout = QGridLayout()
        
        content_layout.addWidget(QLabel("主题:"), 0, 0)
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("可使用变量，如: {date}, {name}")
        content_layout.addWidget(self.subject_edit, 0, 1)
        
        content_layout.addWidget(QLabel("收件人:"), 1, 0)
        self.to_edit = QLineEdit()
        self.to_edit.setPlaceholderText("多个收件人用分号分隔")
        content_layout.addWidget(self.to_edit, 1, 1)
        
        content_layout.addWidget(QLabel("抄送:"), 2, 0)
        self.cc_edit = QLineEdit()
        self.cc_edit.setPlaceholderText("多个抄送用分号分隔")
        content_layout.addWidget(self.cc_edit, 2, 1)
        
        content_layout.addWidget(QLabel("正文:"), 3, 0)
        self.body_edit = QTextEdit()
        self.body_edit.setPlaceholderText("可使用HTML格式和变量，如: {content}, {date}")
        self.body_edit.setAcceptRichText(True)
        content_layout.addWidget(self.body_edit, 3, 1)
        
        content_group.setLayout(content_layout)
        layout.addWidget(content_group)
        
        # 变量说明
        var_group = QGroupBox("可用变量")
        var_layout = QVBoxLayout()
        var_text = QTextEdit()
        var_text.setReadOnly(True)
        var_text.setMaximumHeight(100)
        var_text.setPlainText("""{date} - 当前日期
{name} - 收件人姓名
{content} - 自定义内容
{meeting_title} - 会议标题
{project_name} - 项目名称
{location} - 地点
{time} - 时间
{agenda} - 议程
{notification} - 通知内容""")
        var_layout.addWidget(var_text)
        var_group.setLayout(var_layout)
        layout.addWidget(var_group)
        
        # 操作按钮
        btn_layout = QHBoxLayout()
        self.save_template_btn = QPushButton("保存模板")
        self.save_template_btn.clicked.connect(self.save_template)
        self.cancel_edit_btn = QPushButton("取消")
        self.cancel_edit_btn.clicked.connect(self.cancel_edit)
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_template_btn)
        btn_layout.addWidget(self.cancel_edit_btn)
        layout.addLayout(btn_layout)
        
        self.template_edit_tab.setLayout(layout)
    
    def init_email_generate_tab(self):
        """初始化邮件生成选项卡"""
        layout = QVBoxLayout()
        
        # 模板选择
        template_group = QGroupBox("选择模板")
        template_layout = QHBoxLayout()
        
        template_layout.addWidget(QLabel("邮件模板:"))
        self.email_template_combo = QComboBox()
        self.email_template_combo.currentTextChanged.connect(self.on_email_template_selected)
        template_layout.addWidget(self.email_template_combo)
        
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)
        
        # 邮件参数
        param_group = QGroupBox("邮件参数")
        param_layout = QGridLayout()
        
        param_layout.addWidget(QLabel("主题:"), 0, 0)
        self.email_subject_edit = QLineEdit()
        param_layout.addWidget(self.email_subject_edit, 0, 1)
        
        param_layout.addWidget(QLabel("收件人:"), 1, 0)
        self.email_to_edit = QLineEdit()
        param_layout.addWidget(self.email_to_edit, 1, 1)
        
        param_layout.addWidget(QLabel("抄送:"), 2, 0)
        self.email_cc_edit = QLineEdit()
        param_layout.addWidget(self.email_cc_edit, 2, 1)
        
        param_layout.addWidget(QLabel("正文:"), 3, 0)
        self.email_body_edit = QTextEdit()
        self.email_body_edit.setAcceptRichText(True)
        param_layout.addWidget(self.email_body_edit, 3, 1)
        
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)
        
        # 操作按钮
        btn_layout = QHBoxLayout()
        self.generate_draft_btn = QPushButton("生成邮件草稿")
        self.generate_draft_btn.clicked.connect(self.generate_email_draft)
        self.clear_form_btn = QPushButton("清空表单")
        self.clear_form_btn.clicked.connect(self.clear_email_form)
        btn_layout.addStretch()
        btn_layout.addWidget(self.generate_draft_btn)
        btn_layout.addWidget(self.clear_form_btn)
        layout.addLayout(btn_layout)
        
        self.email_generate_tab.setLayout(layout)
    
    def load_templates(self):
        """加载模板列表"""
        self.template_list.clear()
        templates = self.template_manager.get_all_templates()
        
        for name, template in templates.items():
            item = QListWidgetItem(f"{name} ({template.get('category', '未分类')})")
            item.setData(Qt.UserRole, name)
            self.template_list.addItem(item)
        
        # 更新下拉框
        self.email_template_combo.clear()
        self.email_template_combo.addItems(list(templates.keys()))
    
    def filter_templates(self, category):
        """根据分类过滤模板"""
        self.template_list.clear()
        
        if category == "全部":
            templates = self.template_manager.get_all_templates()
        else:
            templates = self.template_manager.get_templates_by_category(category)
        
        for name, template in templates.items():
            item = QListWidgetItem(f"{name} ({template.get('category', '未分类')})")
            item.setData(Qt.UserRole, name)
            self.template_list.addItem(item)
    
    def on_template_selected(self, item):
        """当选择模板时"""
        # 简化选择模板功能，不再显示预览
        pass
    
    def new_template(self):
        """新建模板"""
        self.tabs.setCurrentIndex(1)  # 切换到编辑选项卡
        self.template_name_edit.clear()
        self.subject_edit.clear()
        self.to_edit.clear()
        self.cc_edit.clear()
        self.body_edit.clear()
        self.template_category_combo.setCurrentIndex(0)
    
    def edit_template(self):
        """编辑模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择一个模板")
            return
        
        template_name = current_item.data(Qt.UserRole)
        template = self.template_manager.get_template(template_name)
        
        if template:
            self.tabs.setCurrentIndex(1)  # 切换到编辑选项卡
            self.template_name_edit.setText(template_name)
            self.subject_edit.setText(template.get('subject', ''))
            self.to_edit.setText(template.get('to', ''))
            self.cc_edit.setText(template.get('cc', ''))
            self.body_edit.setHtml(template.get('body', ''))
            
            category = template.get('category', '工作')
            index = self.template_category_combo.findText(category)
            if index >= 0:
                self.template_category_combo.setCurrentIndex(index)
    
    def delete_template(self):
        """删除模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择一个模板")
            return
        
        template_name = current_item.data(Qt.UserRole)
        reply = QMessageBox.question(self, "确认", f"确定要删除模板 '{template_name}' 吗?",
                                    QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            if self.template_manager.delete_template(template_name):
                QMessageBox.information(self, "成功", "模板已删除")
                self.load_templates()
            else:
                QMessageBox.warning(self, "错误", "删除模板失败")
    
    def save_template(self):
        """保存模板"""
        template_name = self.template_name_edit.text().strip()
        if not template_name:
            QMessageBox.warning(self, "警告", "请输入模板名称")
            return
        
        template_data = {
            "subject": self.subject_edit.text(),
            "to": self.to_edit.text(),
            "cc": self.cc_edit.text(),
            "body": self.body_edit.toHtml(),
            "category": self.template_category_combo.currentText()
        }
        
        if self.template_manager.add_template(template_name, template_data):
            QMessageBox.information(self, "成功", "模板已保存")
            self.load_templates()
            self.tabs.setCurrentIndex(0)  # 切换回模板列表
        else:
            QMessageBox.warning(self, "错误", "保存模板失败")
    
    def cancel_edit(self):
        """取消编辑"""
        self.tabs.setCurrentIndex(0)  # 切换回模板列表
    
    def use_template(self):
        """使用模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择一个模板")
            return
        
        template_name = current_item.data(Qt.UserRole)
        self.tabs.setCurrentIndex(2)  # 切换到邮件生成选项卡
        
        # 设置选中的模板
        index = self.email_template_combo.findText(template_name)
        if index >= 0:
            self.email_template_combo.setCurrentIndex(index)
    
    def on_email_template_selected(self, template_name):
        """当选择邮件模板时"""
        if not template_name:
            return
        
        template = self.template_manager.get_template(template_name)
        if template:
            self.email_subject_edit.setText(template.get('subject', ''))
            self.email_to_edit.setText(template.get('to', ''))
            self.email_cc_edit.setText(template.get('cc', ''))
            self.email_body_edit.setHtml(template.get('body', ''))
    
    def generate_email_draft(self):
        """生成邮件草稿 - 重构版本"""
        
        # 获取邮件数据
        subject = self.email_subject_edit.text().strip()
        to = self.email_to_edit.text().strip()
        cc = self.email_cc_edit.text().strip()
        body = self.email_body_edit.toHtml()
        
        # 验证必填字段
        if not subject:
            QMessageBox.warning(self, "提示", "请输入邮件主题")
            return
        
        if not to:
            QMessageBox.warning(self, "提示", "请输入收件人")
            return
        
        # 替换变量
        from datetime import datetime
        current_date = datetime.now().strftime("%Y年%m月%d日")
        subject = subject.replace("{date}", current_date)
        body = body.replace("{date}", current_date)
        
        # 准备邮件数据
        email_data = {
            'subject': subject,
            'to': to,
            'cc': cc,
            'body': body
        }
        
        # 显示处理状态
        self.generate_draft_btn.setEnabled(False)
        self.generate_draft_btn.setText("处理中...")
        
        # 创建后台线程
        self.draft_thread = EmailDraftGenerator(email_data)
        self.draft_thread.success_signal.connect(self.on_draft_success)
        self.draft_thread.error_signal.connect(self.on_draft_error)
        self.draft_thread.finished.connect(self.on_draft_finished)
        self.draft_thread.start()
    
    def on_draft_success(self, message):
        """邮件草稿生成成功"""
        QMessageBox.information(self, "成功", message)
    
    def on_draft_error(self, error_message):
        """邮件草稿生成失败"""
        QMessageBox.warning(self, "错误", error_message)
    
    def on_draft_finished(self):
        """邮件草稿处理完成"""
        self.generate_draft_btn.setEnabled(True)
        self.generate_draft_btn.setText("生成邮件草稿")
    
    def clear_email_form(self):
        """清空邮件表单"""
        self.email_subject_edit.clear()
        self.email_to_edit.clear()
        self.email_cc_edit.clear()
        self.email_body_edit.clear()