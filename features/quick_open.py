# features/quick_open.py
import os
import json
import subprocess
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QFileDialog, QMessageBox, QMenu,
    QGridLayout, QLabel, QInputDialog
)
from PyQt5.QtCore import Qt

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "quick_open.json")

class QuickOpenWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.data = {}
        self.load_data()
        self.initUI()

    def initUI(self):
        layout = QGridLayout(self)
        layout.setSpacing(12)

        self.buttons = []
        self.path_labels = []

        for i in range(6):
            # 创建按钮
            btn_name = self.data.get(f"btn_{i}", {}).get("name", f"快速打开 {i+1}")
            btn = QPushButton(btn_name)
            btn.setMinimumHeight(40)
            btn.clicked.connect(lambda checked, idx=i: self.open_file(idx))
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, idx=i: self.show_context_menu(idx, pos))
            self.buttons.append(btn)
            layout.addWidget(btn, i, 0)

            # 路径显示
            path_text = self.data.get(f"btn_{i}", {}).get("path", "未选择文件")
            path_label = QLabel(path_text)
            path_label.setStyleSheet("color: #AAAAAA; font-size: 12px;")
            path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.path_labels.append(path_label)
            layout.addWidget(path_label, i, 1)

    def load_data(self):
        """加载配置"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except:
                self.data = {}

    def save_data(self):
        """保存配置"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "保存失败", f"保存路径失败：{e}")

    def show_context_menu(self, idx, pos):
        """右键菜单"""
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
        text, ok = QInputDialog.getText(self, "更改名称", "请输入按钮名称：", text=self.buttons[idx].text())
        if ok and text.strip():
            self.buttons[idx].setText(text.strip())
            key = f"btn_{idx}"
            if key not in self.data:
                self.data[key] = {}
            self.data[key]["name"] = text.strip()
            self.save_data()

    def select_file(self, idx):
        """选择文件路径"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件")
        if file_path:
            key = f"btn_{idx}"
            if key not in self.data:
                self.data[key] = {}
            self.data[key]["path"] = file_path
            self.path_labels[idx].setText(file_path)
            self.save_data()

    def clear_button(self, idx):
        """清除按钮设置"""
        key = f"btn_{idx}"
        if key in self.data:
            del self.data[key]
        self.buttons[idx].setText(f"快速打开 {idx+1}")
        self.path_labels[idx].setText("未选择文件")
        self.save_data()

    def open_file(self, idx):
        """打开文件"""
        key = f"btn_{idx}"
        path = self.data.get(key, {}).get("path", "")
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
