# features/access_database_query.py
import os
import json
import pyodbc
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, 
                            QLabel, QLineEdit, QPushButton, QComboBox, 
                            QTableWidget, QTableWidgetItem, QFileDialog,
                            QMessageBox, QProgressBar, QTextEdit, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from utils.config_manager import load_config, save_config, get_config_value, set_config_value


class DatabaseQueryThread(QThread):
    """数据库查询线程"""
    progress_updated = pyqtSignal(int)
    query_complete = pyqtSignal(list, list)  # 数据, 列名
    query_error = pyqtSignal(str)
    
    def __init__(self, db_path, table_name, columns=None, where_clause=""):
        super().__init__()
        self.db_path = db_path
        self.table_name = table_name
        self.columns = columns or ["*"]
        self.where_clause = where_clause
        self.is_running = False
        
    def run(self):
        """执行查询"""
        self.is_running = True
        try:
            # 构建连接字符串
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                fr'DBQ={self.db_path};'
            )
            
            # 连接数据库
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            # 构建SQL查询
            columns_str = ", ".join(self.columns)
            # 使用方括号包围表名，处理包含空格或特殊字符的表名
            sql = f"SELECT {columns_str} FROM [{self.table_name}]"
            
            if self.where_clause.strip():
                sql += f" WHERE {self.where_clause}"
                
            # 执行查询
            self.progress_updated.emit(30)
            cursor.execute(sql)
            
            # 获取列名
            column_names = [column[0] for column in cursor.description]
            
            # 获取数据
            data = []
            rows = cursor.fetchall()
            total_rows = len(rows)
            
            for i, row in enumerate(rows):
                if not self.is_running:
                    break
                    
                data_row = []
                for value in row:
                    # 处理二进制数据
                    if isinstance(value, bytes):
                        data_row.append("<二进制数据>")
                    else:
                        data_row.append(str(value) if value is not None else "")
                
                data.append(data_row)
                
                # 更新进度
                progress = 30 + int((i + 1) / total_rows * 60)
                self.progress_updated.emit(progress)
            
            # 关闭连接
            conn.close()
            
            if self.is_running:
                self.progress_updated.emit(100)
                self.query_complete.emit(data, column_names)
                
        except Exception as e:
            if self.is_running:
                self.query_error.emit(f"查询错误: {str(e)}")
                
    def stop(self):
        """停止查询"""
        self.is_running = False


class AccessDatabaseQueryWidget(QWidget):
    """Access数据库查询工具"""
    
    def __init__(self):
        super().__init__()
        self.db_connection = None
        self.current_db_path = ""
        self.query_thread = None
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        """初始化UI界面"""
        layout = QVBoxLayout(self)
        
        # 创建分割器
        splitter = QSplitter(Qt.Vertical)
        
        # 上半部分：控制面板
        control_panel = QWidget()
        control_layout = QVBoxLayout(control_panel)
        
        # 数据库连接配置
        db_group = QGroupBox("数据库连接")
        db_layout = QVBoxLayout()
        
        # 数据库文件选择
        db_file_layout = QHBoxLayout()
        db_file_layout.addWidget(QLabel("Access数据库:"))
        self.db_path_input = QLineEdit()
        self.db_path_input.setPlaceholderText("选择Access数据库文件(.mdb或.accdb)...")
        self.db_path_input.textChanged.connect(self.on_db_path_changed)
        db_browse_btn = QPushButton("浏览...")
        db_browse_btn.clicked.connect(self.browse_database)
        db_connect_btn = QPushButton("连接")
        db_connect_btn.clicked.connect(self.connect_database)
        db_file_layout.addWidget(self.db_path_input)
        db_file_layout.addWidget(db_browse_btn)
        db_file_layout.addWidget(db_connect_btn)
        db_layout.addLayout(db_file_layout)
        
        # 连接状态
        self.connection_status = QLabel("未连接")
        self.connection_status.setStyleSheet("color: red; font-weight: bold;")
        db_layout.addWidget(self.connection_status)
        
        db_group.setLayout(db_layout)
        control_layout.addWidget(db_group)
        
        # 查询配置
        query_group = QGroupBox("查询配置")
        query_layout = QVBoxLayout()
        
        # 表名选择
        table_layout = QHBoxLayout()
        table_layout.addWidget(QLabel("表名:"))
        self.table_combo = QComboBox()
        self.table_combo.currentTextChanged.connect(self.on_table_changed)
        table_layout.addWidget(self.table_combo)
        query_layout.addLayout(table_layout)
        
        # 列名选择
        columns_layout = QHBoxLayout()
        columns_layout.addWidget(QLabel("列名:"))
        self.columns_combo = QComboBox()
        self.columns_combo.addItem("*")  # 默认添加所有列选项
        columns_layout.addWidget(self.columns_combo)
        query_layout.addLayout(columns_layout)
        
        # 查询条件
        where_layout = QHBoxLayout()
        where_layout.addWidget(QLabel("搜索内容:"))
        self.where_input = QLineEdit()
        self.where_input.setPlaceholderText("输入要搜索的内容，按Enter键执行搜索")
        # 添加回车键触发搜索
        self.where_input.returnPressed.connect(self.execute_query)
        where_layout.addWidget(self.where_input)
        query_layout.addLayout(where_layout)
        
        # 查询按钮
        btn_layout = QHBoxLayout()
        self.query_btn = QPushButton("执行查询")
        self.query_btn.clicked.connect(self.execute_query)
        self.query_btn.setEnabled(False)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_query)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.query_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()
        query_layout.addLayout(btn_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        query_layout.addWidget(self.progress_bar)
        
        query_group.setLayout(query_layout)
        control_layout.addWidget(query_group)
        
        # 统计信息
        self.stats_label = QLabel("准备就绪")
        self.stats_label.setStyleSheet("font-weight: bold; color: #666;")
        control_layout.addWidget(self.stats_label)
        
        control_layout.addStretch()
        
        # 下半部分：结果显示
        result_widget = QWidget()
        result_layout = QVBoxLayout(result_widget)
        
        # 结果表格
        self.result_table = QTableWidget()
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setSelectionBehavior(QTableWidget.SelectItems)  # 允许选择单个单元格
        self.result_table.setSelectionMode(QTableWidget.SingleSelection)  # 单选模式
        self.result_table.setAlternatingRowColors(True)
        result_layout.addWidget(self.result_table)
        
        # 添加到分割器
        splitter.addWidget(control_panel)
        splitter.addWidget(result_widget)
        
        # 设置分割器比例
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        
        # 添加到主布局
        layout.addWidget(splitter)
        
    def browse_database(self):
        """浏览选择数据库文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Access数据库文件", "", 
            "Access数据库文件 (*.mdb *.accdb);;所有文件 (*)"
        )
        
        if file_path:
            self.db_path_input.setText(file_path)
            
    def on_db_path_changed(self):
        """数据库路径改变时保存设置"""
        db_path = self.db_path_input.text().strip()
        if db_path:
            set_config_value("access_database_query", {"db_path": db_path})
            
    def connect_database(self):
        """连接到Access数据库"""
        db_path = self.db_path_input.text().strip()
        
        if not db_path:
            QMessageBox.warning(self, "警告", "请选择Access数据库文件")
            return
            
        if not os.path.exists(db_path):
            QMessageBox.warning(self, "警告", "文件不存在")
            return
            
        try:
            # 构建连接字符串
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                fr'DBQ={db_path};'
            )
            
            # 连接数据库
            self.db_connection = pyodbc.connect(conn_str)
            self.current_db_path = db_path
            
            # 更新连接状态
            self.connection_status.setText("已连接")
            self.connection_status.setStyleSheet("color: green; font-weight: bold;")
            
            # 获取表列表
            cursor = self.db_connection.cursor()
            cursor.tables()
            tables = []
            
            # 获取所有表名，包括系统表
            for row in cursor.fetchall():
                table_name = row.table_name
                # 过滤掉系统表（以MSys开头的表）
                if not table_name.startswith("MSys") and row.table_type == "TABLE":
                    tables.append(table_name)
            
            # 如果没有找到表，尝试使用另一种方法
            if not tables:
                try:
                    # 直接查询系统表获取表名
                    cursor.execute("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
                    tables = [row[0] for row in cursor.fetchall()]
                except:
                    pass
            
            # 更新表名下拉框
            self.table_combo.clear()
            self.table_combo.addItems(tables)
            
            # 启用查询按钮
            self.query_btn.setEnabled(True)
            
            # 保存设置
            set_config_value("access_database_query", {"db_path": db_path})
            
            QMessageBox.information(self, "成功", "数据库连接成功")
            
        except Exception as e:
            self.connection_status.setText("连接失败")
            self.connection_status.setStyleSheet("color: red; font-weight: bold;")
            QMessageBox.critical(self, "错误", f"连接数据库失败: {str(e)}")
            
    def on_table_changed(self):
        """表名改变时获取列信息"""
        if not self.db_connection or not self.table_combo.currentText():
            return
            
        try:
            cursor = self.db_connection.cursor()
            table_name = self.table_combo.currentText()
            
            # 获取列信息
            cursor.columns(table=table_name)
            columns = [row.column_name for row in cursor.fetchall()]
            
            # 如果没有获取到列信息，尝试直接查询表结构
            if not columns:
                try:
                    # 使用方括号包围表名，处理包含空格或特殊字符的表名
                    cursor.execute(f"SELECT TOP 1 * FROM [{table_name}]")
                    columns = [column[0] for column in cursor.description]
                except Exception as e:
                    print(f"直接查询表结构失败: {str(e)}")
                    return
            
            # 更新列名下拉框
            self.columns_combo.clear()
            self.columns_combo.addItem("*")  # 添加所有列选项
            self.columns_combo.addItems(columns)  # 添加具体列名
            
        except Exception as e:
            print(f"获取列信息失败: {str(e)}")
            
    def execute_query(self):
        """执行查询"""
        if not self.db_connection:
            QMessageBox.warning(self, "警告", "请先连接数据库")
            return
            
        table_name = self.table_combo.currentText()
        if not table_name:
            QMessageBox.warning(self, "警告", "请选择表名")
            return
            
        # 获取列名
        selected_column = self.columns_combo.currentText()
        
        # 无论选择哪一列，都查询所有列，这样可以看到完整记录
        columns = ["*"]
        
        # 获取搜索内容
        search_text = self.where_input.text().strip()
        where_clause = ""
        
        # 如果用户输入了搜索内容，构建WHERE子句
        if search_text:
            if selected_column == "*":
                # 如果选择了所有列，则在所有可搜索的列中搜索
                try:
                    cursor = self.db_connection.cursor()
                    # 获取表的所有列信息
                    cursor.columns(table=table_name)
                    column_info = cursor.fetchall()
                    
                    # 筛选出所有可能包含文本的列（包括数字列，因为Access可以隐式转换）
                    searchable_columns = []
                    for col in column_info:
                        # 排除二进制类型的列，其他类型都尝试搜索
                        if col.type_name not in ('BINARY', 'VARBINARY', 'LONGVARBINARY', 'IMAGE', 'OLEOBJECT'):
                            searchable_columns.append(f"[{col.column_name}]")
                    
                    if searchable_columns:
                        # 构建OR条件，在所有可搜索的列中查找匹配的内容
                        conditions = [f"{col} LIKE '%{search_text}%'" for col in searchable_columns]
                        where_clause = " OR ".join(conditions)
                except Exception as e:
                    print(f"获取列信息失败: {str(e)}")
                    where_clause = f"1=1"  # 如果无法获取列信息，则查询所有记录
            else:
                # 如果选择了特定列，则在该列中搜索
                where_clause = f"[{selected_column}] LIKE '%{search_text}%'"
        
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # 禁用查询按钮，启用停止按钮
        self.query_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # 创建并启动查询线程
        self.query_thread = DatabaseQueryThread(
            self.current_db_path, table_name, columns, where_clause
        )
        self.query_thread.progress_updated.connect(self.update_progress)
        self.query_thread.query_complete.connect(self.display_results)
        self.query_thread.query_error.connect(self.on_query_error)
        self.query_thread.start()
        
    def stop_query(self):
        """停止查询"""
        if self.query_thread and self.query_thread.isRunning():
            self.query_thread.stop()
            self.query_thread.wait()
            
        self.progress_bar.setVisible(False)
        self.query_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.stats_label.setText("查询已停止")
        
    def update_progress(self, value):
        """更新进度"""
        self.progress_bar.setValue(value)
        
    def display_results(self, data, column_names):
        """显示查询结果"""
        # 隐藏进度条
        self.progress_bar.setVisible(False)
        
        # 恢复按钮状态
        self.query_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # 设置表格
        self.result_table.setRowCount(len(data))
        self.result_table.setColumnCount(len(column_names))
        self.result_table.setHorizontalHeaderLabels(column_names)
        
        # 填充数据
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                self.result_table.setItem(row_idx, col_idx, item)
                
        # 调整列宽
        self.result_table.resizeColumnsToContents()
        
        # 更新统计信息
        self.stats_label.setText(f"查询完成，共 {len(data)} 条记录")
        
    def on_query_error(self, error_msg):
        """查询错误处理"""
        # 隐藏进度条
        self.progress_bar.setVisible(False)
        
        # 恢复按钮状态
        self.query_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # 显示错误
        QMessageBox.critical(self, "查询错误", error_msg)
        self.stats_label.setText(f"查询失败: {error_msg}")
        
    def load_settings(self):
        """加载设置"""
        config = get_config_value("access_database_query", {})
        db_path = config.get("db_path", "")
        
        if db_path and os.path.exists(db_path):
            self.db_path_input.setText(db_path)
            
    def closeEvent(self, event):
        """关闭事件"""
        if self.db_connection:
            self.db_connection.close()
        event.accept()