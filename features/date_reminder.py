# features/date_reminder.py
import os
import json
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, 
                            QTableWidgetItem, QPushButton, QDialog, QLabel, 
                            QLineEdit, QDateEdit, QComboBox, QMessageBox, 
                            QHeaderView)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor


DATA_FILE = "dates.json"


class DateReminderWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.load_reminders()



    def init_ui(self):
        """初始化UI - 恢复原始风格"""
        self.setStyleSheet("""
            QWidget {
                background-color: #F8F9FA;
                color: #4A5568;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
            }
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #F1F5F9;
                gridline-color: #CBD5E0;
                border: 2px solid #CBD5E0;
                border-radius: 8px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
                font-size: 18px;
            }
            QTableWidget::item {
                padding: 12px;
                color: #2D3748;
                font-weight: 500;
            }
            QHeaderView::section {
                background-color: #4A5568;
                color: white;
                padding: 15px 10px;
                font-weight: bold;
                font-size: 18px;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
            }
            QPushButton {
                background-color: #4A5568;
                color: white;
                border: none;
                padding: 15px 25px;
                border-radius: 8px;
                font-size: 18px;
                font-weight: bold;
                font-family: 'Segoe UI', 'Microsoft YaHei', '等线';
            }
            QPushButton:hover {
                background-color: #2D3748;
            }
            QPushButton:pressed {
                background-color: #1A202C;
            }
            QTableWidget::item:selected {
                background-color: #E2E8F0;
                color: #1A202C;
            }
            QTableWidget::item:focus {
                outline: none;
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
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 表格 - 4列版本，保持原始风格
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["事件名称", "日期", "剩余天数", "备注"])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # 设置表格列宽比例：事件名称9/20，日期3/20，剩余天数3/20，备注5/20
        total_width = 1236

# 设置表格列宽比例：事件名称9/20，日期3/20，剩余天数3/20，备注5/20（总宽度800px）
        self.table.setColumnWidth(0, 559)  # 事件名称 9/20 × 800 = 360px
        self.table.setColumnWidth(1, 184)  # 日期 3/20 × 800 = 120px
        self.table.setColumnWidth(2, 184)  # 剩余天数 3/20 × 800 = 120px
        self.table.setColumnWidth(3, 309)  # 备注 5/20 × 800 = 200px
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        layout.addWidget(self.table)

        # 按钮 - 恢复原始文本和风格
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("➕ 添加事件")
        self.btn_edit = QPushButton("✏ 编辑事件")
        self.btn_delete = QPushButton("🗑 删除事件")

        self.btn_add.clicked.connect(self.add_reminder)
        self.btn_edit.clicked.connect(self.edit_reminder)
        self.btn_delete.clicked.connect(self.delete_reminder)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)



    def add_reminder(self):
        """添加新提醒"""
        dialog = QDialog(self)
        dialog.setWindowTitle("添加提醒")
        dialog.setModal(True)
        dialog.resize(400, 300)

        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 事件名称
        name_label = QLabel("事件名称:")
        name_input = QLineEdit()
        name_input.setPlaceholderText("请输入事件名称")
        name_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 日期
        date_label = QLabel("日期:")
        date_input = QDateEdit()
        date_input.setCalendarPopup(True)
        date_input.setDate(QDate.currentDate())
        date_input.setDisplayFormat("yyyy-MM-dd")
        date_input.setStyleSheet("""
            QDateEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 备注
        note_label = QLabel("备注:")
        note_input = QLineEdit()
        note_input.setPlaceholderText("请输入备注信息")
        note_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 颜色选择
        color_label = QLabel("颜色标记:")
        color_combo = QComboBox()
        color_combo.addItems(["默认", "红色", "蓝色", "绿色", "黄色", "紫色"])
        color_combo.setStyleSheet("""
            QComboBox {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        layout.addWidget(name_label)
        layout.addWidget(name_input)
        layout.addWidget(date_label)
        layout.addWidget(date_input)
        layout.addWidget(note_label)
        layout.addWidget(note_input)
        layout.addWidget(color_label)
        layout.addWidget(color_combo)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_cancel = QPushButton("取消")
        
        btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #4A5568;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 16px;
            }
        """)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #CBD5E0;
                color: #4A5568;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 16px;
            }
        """)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        btn_ok.clicked.connect(dialog.accept)
        btn_cancel.clicked.connect(dialog.reject)

        if dialog.exec_() == QDialog.Accepted:
            if not name_input.text().strip():
                QMessageBox.warning(self, "警告", "请输入事件名称！")
                return

            reminder = {
                'name': name_input.text().strip(),
                'date': date_input.date().toString("yyyy-MM-dd"),
                'note': note_input.text().strip(),
                'color': color_combo.currentText()
            }

            self.reminders.append(reminder)
            self.save_reminders()
            self.refresh_table()
            QMessageBox.information(self, "成功", "提醒添加成功！")

    def save_reminders(self):
        """保存提醒数据"""
        os.makedirs('data', exist_ok=True)
        with open('data/reminders.json', 'w', encoding='utf-8') as f:
            json.dump(self.reminders, f, ensure_ascii=False, indent=2)

    def delete_reminder(self):
        """删除选中的提醒"""
        current_row = self.table.currentRow()
        if current_row == -1:
            QMessageBox.warning(self, "警告", "请先选择要删除的提醒！")
            return

        # 获取表格中选中行的事件名称和日期，用于在原始列表中找到对应项
        selected_name = self.table.item(current_row, 0).text()
        selected_date = self.table.item(current_row, 1).text()
        
        # 在原始列表中查找匹配的项
        found_index = -1
        for i, reminder in enumerate(self.reminders):
            if reminder['name'] == selected_name and reminder['date'] == selected_date:
                found_index = i
                break
        
        if found_index == -1:
            QMessageBox.warning(self, "错误", "无法找到选中的提醒！")
            return

        reply = QMessageBox.question(
            self, "确认删除", "确定要删除这个提醒吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            del self.reminders[found_index]
            self.save_reminders()
            self.refresh_table()
            QMessageBox.information(self, "成功", "提醒已删除！")

    def edit_reminder(self):
        """编辑选中的提醒"""
        current_row = self.table.currentRow()
        if current_row == -1:
            QMessageBox.warning(self, "警告", "请先选择要编辑的提醒！")
            return

        reminder = self.reminders[current_row]
        
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑提醒")
        dialog.setModal(True)
        dialog.resize(400, 300)

        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 事件名称
        name_label = QLabel("事件名称:")
        name_input = QLineEdit(reminder['name'])
        name_input.setPlaceholderText("请输入事件名称")
        name_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 日期
        date_label = QLabel("日期:")
        date_input = QDateEdit()
        date_input.setCalendarPopup(True)
        date_input.setDate(QDate.fromString(reminder['date'], "yyyy-MM-dd"))
        date_input.setDisplayFormat("yyyy-MM-dd")
        date_input.setStyleSheet("""
            QDateEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 备注
        note_label = QLabel("备注:")
        note_input = QLineEdit(reminder.get('note', ''))
        note_input.setPlaceholderText("请输入备注信息")
        note_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        # 颜色选择
        color_label = QLabel("颜色标记:")
        color_combo = QComboBox()
        color_combo.addItems(["默认", "红色", "蓝色", "绿色", "黄色", "紫色"])
        color_combo.setCurrentText(reminder.get('color', '默认'))
        color_combo.setStyleSheet("""
            QComboBox {
                padding: 10px;
                border: 1px solid #CBD5E0;
                border-radius: 5px;
                font-size: 16px;
            }
        """)

        layout.addWidget(name_label)
        layout.addWidget(name_input)
        layout.addWidget(date_label)
        layout.addWidget(date_input)
        layout.addWidget(note_label)
        layout.addWidget(note_input)
        layout.addWidget(color_label)
        layout.addWidget(color_combo)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_cancel = QPushButton("取消")
        
        btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #4A5568;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 16px;
            }
        """)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #CBD5E0;
                color: #4A5568;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 16px;
            }
        """)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        btn_ok.clicked.connect(dialog.accept)
        btn_cancel.clicked.connect(dialog.reject)

        if dialog.exec_() == QDialog.Accepted:
            if not name_input.text().strip():
                QMessageBox.warning(self, "警告", "请输入事件名称！")
                return

            reminder['name'] = name_input.text().strip()
            reminder['date'] = date_input.date().toString("yyyy-MM-dd")
            reminder['note'] = note_input.text().strip()
            reminder['color'] = color_combo.currentText()

            self.save_reminders()
            self.refresh_table()
            QMessageBox.information(self, "成功", "提醒修改成功！")

    def load_reminders(self):
        """加载提醒数据"""
        try:
            with open('data/reminders.json', 'r', encoding='utf-8') as f:
                self.reminders = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.reminders = []
        
        self.refresh_table()

    def refresh_table(self):
        """刷新表格显示 - 按到期时间自动排序"""
        self.table.setRowCount(0)
        
        # 按到期时间排序（快到期的排前面）
        sorted_reminders = sorted(self.reminders, key=lambda x: 
            (datetime.strptime(x['date'], "%Y-%m-%d") - datetime.now()).days)
        
        for reminder in sorted_reminders:
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            # 事件名称
            name_item = QTableWidgetItem(reminder['name'])
            name_item.setTextAlignment(Qt.AlignCenter)
            
            # 日期
            date_str = reminder['date']
            date_item = QTableWidgetItem(date_str)
            date_item.setTextAlignment(Qt.AlignCenter)
            
            # 计算剩余天数
            try:
                target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                today = datetime.now().date()
                days_left = (target_date - today).days
            except ValueError:
                days_left = 0
            
            # 剩余天数
            days_item = QTableWidgetItem(str(days_left))
            days_item.setTextAlignment(Qt.AlignCenter)
            
            # 备注
            note_item = QTableWidgetItem(reminder.get('note', ''))
            note_item.setTextAlignment(Qt.AlignCenter)
            
            # 设置颜色标记和橙色高亮
            color = reminder.get('color', '默认')
            color_map = {
                '红色': QColor('#FF0000'),
                '蓝色': QColor('#0000FF'),
                '绿色': QColor('#008000'),
                '黄色': QColor('#FFA500'),
                '紫色': QColor('#800080'),
                '默认': QColor('#2D3748')
            }
            
            text_color = color_map.get(color, QColor('#2D3748'))
            
            # 剩余天数≤1时橙色高亮（包括等于1和小于1）
            if days_left <= 1:
                orange_bg = QColor('#FFF3E0')
                orange_text = QColor('#FF6B35')
                
                name_item.setBackground(orange_bg)
                name_item.setForeground(orange_text)
                date_item.setBackground(orange_bg)
                date_item.setForeground(orange_text)
                days_item.setBackground(orange_bg)
                days_item.setForeground(orange_text)
                note_item.setBackground(orange_bg)
                note_item.setForeground(orange_text)
            else:
                # 应用颜色标记
                name_item.setForeground(text_color)
                date_item.setForeground(text_color)
                days_item.setForeground(text_color)
                note_item.setForeground(text_color)
            
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, date_item)
            self.table.setItem(row, 2, days_item)
            self.table.setItem(row, 3, note_item)
