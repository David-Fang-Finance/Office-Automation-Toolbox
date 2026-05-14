# main.py
from PyQt5.QtWidgets import (
    QApplication, QWidget, QTreeWidget, QTreeWidgetItem,
    QStackedWidget, QHBoxLayout, QPushButton, QLabel, QGridLayout
)
from PyQt5.QtGui import QPainter, QColor, QLinearGradient, QBrush, QFont
from PyQt5.QtCore import Qt, QPropertyAnimation, pyqtProperty, QRectF, QEasingCurve
from features.sap_integrated import SapIntegratedWidget
from features.pdf_split import PdfSplitWidget
from features.date_reminder import DateReminderWidget
from features.quick_open import QuickOpenWidget
from features.invoice_extractor import InvoiceExtractorWidget
from features.excel_merger import ExcelMergerWidget
from features.ocr_invoice_extractor import OCRInvoiceExtractorWidget
from features.gemini_invoice_extractor import GeminiInvoiceExtractorWidget
from features.access_database_query import AccessDatabaseQueryWidget
from features.email_automation import EmailAutomationWidget

# ==========================================
# 科技按钮样式 - 来自确定按钮.txt
# ==========================================
class TechBarButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)
        
        # 保持原长度，不改变按钮的大小
        self.setMinimumHeight(40)
        
        # 动画进度 (0.0 -> 1.0)
        self._anim_progress = 0.0
        
        # 定义动画：光流划过
        self.anim = QPropertyAnimation(self, b"progress")
        self.anim.setDuration(250) # 0.25秒，极速响应
        self.anim.setEasingCurve(QEasingCurve.OutQuad)

    # --- 动画属性钩子 ---
    @pyqtProperty(float)
    def progress(self): return self._anim_progress
    @progress.setter
    def progress(self, val):
        self._anim_progress = val
        self.update() # 刷新重绘

    # --- 鼠标交互 ---
    def enterEvent(self, e):
        # 鼠标悬停：进度条 0 -> 1
        self.anim.stop()
        self.anim.setStartValue(self._anim_progress)
        self.anim.setEndValue(1.0)
        self.anim.start()
        super().enterEvent(e)

    def leaveEvent(self, e):
        # 鼠标离开：进度条 1 -> 0
        self.anim.stop()
        self.anim.setStartValue(self._anim_progress)
        self.anim.setEndValue(0.0)
        self.anim.start()
        super().leaveEvent(e)

    # --- 核心绘图 (只改这里，让它变帅) ---
    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        w, h = self.width(), self.height()

        # 1. 绘制底色 (你原来的深灰色)
        # 颜色代码 #4B5563 是类似你截图的深灰蓝
        bg_color = QColor("#4B5563") 
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(0, 0, w, h, 6, 6) # 圆角 6px

        # 2. 绘制光流悬停效果 (Wipe Effect)
        if self._anim_progress > 0.01:
            # 渐变色：青色 -> 蓝色 (赛博朋克感)
            gradient = QLinearGradient(0, 0, w, 0)
            gradient.setColorAt(0, QColor("#06b6d4")) # Cyan
            gradient.setColorAt(1, QColor("#3b82f6")) # Blue
            
            painter.setBrush(QBrush(gradient))
            
            # 计算光流覆盖的宽度
            flow_width = w * self._anim_progress
            
            # 绘制覆盖层 (用 Clip 限制圆角)
            painter.save()
            painter.setClipRect(0, 0, int(flow_width), h)
            painter.drawRoundedRect(0, 0, w, h, 6, 6)
            painter.restore()

        # 3. 绘制文字 (居中)
        painter.setPen(Qt.white)
        painter.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        painter.drawText(self.rect(), Qt.AlignCenter, self.text())

        # 4. 绘制边框 (悬停时发光)
        # 平时没边框，悬停时有一圈细细的亮边
        if self._anim_progress > 0:
            border_color = QColor(255, 255, 255, int(100 * self._anim_progress))
            painter.setBrush(Qt.NoBrush)
            painter.setPen(border_color)
            painter.drawRoundedRect(1, 1, w-2, h-2, 6, 6)

# ==========================================
# 科技感快速访问控件 - 包装原有的QuickOpenWidget
# ==========================================
class TechQuickOpenWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.original_widget = QuickOpenWidget()
        self.init_ui()

    def init_ui(self):
        layout = QGridLayout(self)
        layout.setSpacing(12)

        # 获取原始控件的数据
        self.original_widget.load_data()
        
        self.buttons = []
        self.path_labels = []

        for i in range(6):
            # 创建科技按钮
            btn_data = self.original_widget.data.get(f"btn_{i}", {})
            btn_name = btn_data.get("name", f"快速打开 {i+1}")
            btn = TechBarButton(btn_name)
            btn.clicked.connect(lambda checked, idx=i: self.open_file(idx))
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, idx=i: self.show_context_menu(idx, pos))
            self.buttons.append(btn)
            layout.addWidget(btn, i, 0)

            # 路径显示
            path_text = btn_data.get("path", "未选择文件")
            path_label = QLabel(path_text)
            path_label.setStyleSheet("color: #AAAAAA; font-size: 12px;")
            path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.path_labels.append(path_label)
            layout.addWidget(path_label, i, 1)

    def load_data(self):
        """加载配置"""
        return self.original_widget.load_data()

    def save_data(self):
        """保存配置"""
        return self.original_widget.save_data()

    def show_context_menu(self, idx, pos):
        """右键菜单"""
        from PyQt5.QtWidgets import QMenu
        menu = QMenu(self)

        action_name = menu.addAction("更改名称")
        action_path = menu.addAction("更改路径")
        action_clear = menu.addAction("清除设置")

        action = menu.exec_(self.buttons[idx].mapToGlobal(pos))

        if action == action_name:
            self.change_name(idx)
        elif action == action_path:
            self.select_file(idx)
        elif action == action_clear:
            self.clear_button(idx)

    def change_name(self, idx):
        """修改按钮名称"""
        from PyQt5.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(self, "更改名称", "请输入按钮名称：", text=self.buttons[idx].text())
        if ok and text.strip():
            self.buttons[idx].setText(text.strip())
            key = f"btn_{idx}"
            if key not in self.original_widget.data:
                self.original_widget.data[key] = {}
            self.original_widget.data[key]["name"] = text.strip()
            self.save_data()

    def select_file(self, idx):
        """选择文件路径"""
        from PyQt5.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件")
        if file_path:
            key = f"btn_{idx}"
            if key not in self.original_widget.data:
                self.original_widget.data[key] = {}
            self.original_widget.data[key]["path"] = file_path
            self.path_labels[idx].setText(file_path)
            self.save_data()

    def clear_button(self, idx):
        """清除按钮设置"""
        key = f"btn_{idx}"
        if key in self.original_widget.data:
            del self.original_widget.data[key]
        self.buttons[idx].setText(f"快速打开 {idx+1}")
        self.path_labels[idx].setText("未选择文件")
        self.save_data()

    def open_file(self, idx):
        """打开文件"""
        import os
        import subprocess
        from PyQt5.QtWidgets import QMessageBox
        
        key = f"btn_{idx}"
        path = self.original_widget.data.get(key, {}).get("path", "")
        if path and os.path.exists(path):
            try:
                if os.name == 'nt':
                    os.startfile(path)
                else:
                    subprocess.Popen(["open", path])
            except Exception as e:
                QMessageBox.warning(self, "打开失败", f"无法打开文件：{e}")
        else:
            QMessageBox.information(self, "未设置路径", "请右键按钮设置文件路径")


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("工具箱")
        self.resize(1600, 1050)

        # 高级灰调 + 深度对比
        self.setStyleSheet("""
            QWidget { 
                background-color: #F8F9FA; 
                color: #2D3748; 
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线'; 
                font-size: 18px; 
            }
            QTreeWidget { 
                background-color: #FFFFFF; 
                border: none;
                border-right: 2px solid #CBD5E0;
                font-size: 20px; 
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                outline: none;
            }
            QTreeWidget::item {
                padding: 20px 35px;
                border: none;
                margin: 0px;
                border-bottom: 1px solid #E2E8F0;
                color: #4A5568;
            }
            QTreeWidget::item:selected {
                background-color: #F1F5F9;
                color: #1A202C;
                border-left: 4px solid #4A5568;
                font-weight: 600;
            }
            QTreeWidget::item:hover {
                background-color: #EDF2F7;
                color: #2D3748;
            }
            QTreeWidget::item:selected:hover {
                background-color: #E2E8F0;
                color: #1A202C;
            }
            QPushButton { 
                background-color: #4A5568; 
                border: none; 
                padding: 16px 28px; 
                border-radius: 8px; 
                font-size: 17px; 
                color: #FFFFFF;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-weight: 600;
            }
            QPushButton:hover { 
                background-color: #2D3748; 
            }
            QPushButton:pressed {
                background-color: #1A202C;
            }
            QTextEdit { 
                background-color: #FFFFFF; 
                border: 2px solid #CBD5E0; 
                border-radius: 8px; 
                padding: 18px; 
                font-size: 18px; 
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #2D3748;
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 2px solid #CBD5E0;
                border-radius: 8px;
                padding: 14px;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #2D3748;
            }
            QLineEdit:focus {
                border: 2px solid #4A5568;
            }
            QLabel {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
                color: #4A5568;
                font-weight: 500;
            }
            QGroupBox {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 20px;
                font-weight: 700;
                color: #1A202C;
                border: 2px solid #CBD5E0;
                border-radius: 12px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 20px;
                padding: 0 10px 0 10px;
            }
            QTabWidget::pane {
                border: 2px solid #CBD5E0;
                background-color: #FFFFFF;
                border-radius: 8px;
            }
            QTabWidget::tab-bar {
                alignment: left;
            }
            QTabBar::tab {
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
                padding: 18px 30px;
                background-color: #F8F9FA;
                border: 2px solid #CBD5E0;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 2px;
                color: #718096;
                font-weight: 500;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                border-bottom: 3px solid #4A5568;
                color: #1A202C;
                font-weight: 600;
            }
            QTabBar::tab:hover:!selected {
                background-color: #F1F5F9;
                color: #2D3748;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 2px solid #CBD5E0;
                border-radius: 8px;
                padding: 14px;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #2D3748;
                min-height: 45px;
            }
            QComboBox:hover {
                border: 2px solid #4A5568;
            }
            QComboBox::drop-down {
                border: none;
                width: 40px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 8px solid #4A5568;
                margin-right: 15px;
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
            QSpinBox {
                background-color: #FFFFFF;
                border: 2px solid #CBD5E0;
                border-radius: 8px;
                padding: 14px;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                color: #2D3748;
                font-weight: 500;
            }
            QSpinBox:focus {
                border: 2px solid #4A5568;
            }
            QScrollBar:vertical {
                background-color: #F8F9FA;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #A0AEC0;
                border-radius: 6px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #4A5568;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
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

        # 左侧扁平菜单
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setFixedWidth(320)
        self.tree.setRootIsDecorated(False)

        # 创建扁平菜单项 - 深色调
        menu_items = [
            "⏰ 提醒助手",
            "🎯 SAP工具",
            "📄 PDF处理",
            "📊 Excel工具",
            "⚡ 快速访问",
            "🧾 发票提取",
            "🔍 OCR发票提取",
            "✨ Gemini发票提取",
            "🗃️ Access数据库查询",
            "📧 邮件自动化助手"
        ]

        for item_text in menu_items:
            menu_item = QTreeWidgetItem(self.tree)
            menu_item.setText(0, item_text)

        # 右侧页面
        self.stack = QStackedWidget()
        self.stack.addWidget(DateReminderWidget())
        self.stack.addWidget(SapIntegratedWidget())
        self.stack.addWidget(PdfSplitWidget())
        self.stack.addWidget(ExcelMergerWidget())
        self.stack.addWidget(TechQuickOpenWidget())
        self.stack.addWidget(InvoiceExtractorWidget())
        self.stack.addWidget(OCRInvoiceExtractorWidget())
        self.stack.addWidget(GeminiInvoiceExtractorWidget())
        self.stack.addWidget(AccessDatabaseQueryWidget())
        self.stack.addWidget(EmailAutomationWidget())

        # 布局
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.tree)
        layout.addWidget(self.stack)

        # 点击菜单切换右侧页面
        self.tree.itemClicked.connect(self.on_tree_item_clicked)

    def on_tree_item_clicked(self, item, column):
        text = item.text(0)
        menu_mapping = {
            "⏰ 提醒助手": 0,
            "🎯 SAP工具": 1,
            "📄 PDF处理": 2,
            "📊 Excel工具": 3,
            "⚡ 快速访问": 4,
            "🧾 发票提取": 5,
            "🔍 OCR发票提取": 6,
            "✨ Gemini发票提取": 7,
            "🗃️ Access数据库查询": 8,
            "📧 邮件自动化助手": 9,
        }
        
        if text in menu_mapping:
            self.stack.setCurrentIndex(menu_mapping[text])


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
