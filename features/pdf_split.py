# features/pdf_split.py
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QLineEdit, QMessageBox, QGroupBox, QListWidget, QHBoxLayout
)
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter
import os
from PIL import Image

class PdfSplitWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_files = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.title = QLabel("PDF 处理工具")
        self.title.setObjectName("titleLabel")
        layout.addWidget(self.title)

        # PDF拆分功能组
        split_group = QGroupBox("PDF 拆分")
        split_layout = QVBoxLayout()
        
        # 选择 PDF 文件
        self.btn_select_pdf = QPushButton("选择 PDF 文件")
        self.btn_select_pdf.clicked.connect(self.select_pdf_file)
        split_layout.addWidget(self.btn_select_pdf)

        # 选择 PDF 文件夹
        self.btn_select_folder = QPushButton("选择 PDF 文件夹")
        self.btn_select_folder.clicked.connect(self.select_pdf_folder)
        split_layout.addWidget(self.btn_select_folder)

        # 输出文件夹
        self.line_output = QLineEdit()
        self.line_output.setPlaceholderText("选择输出文件夹")
        split_layout.addWidget(self.line_output)

        self.btn_select_output = QPushButton("选择输出文件夹")
        self.btn_select_output.clicked.connect(self.select_output_folder)
        split_layout.addWidget(self.btn_select_output)

        # 执行拆分
        self.btn_split = QPushButton("开始拆分")
        self.btn_split.clicked.connect(self.split_pdf)
        split_layout.addWidget(self.btn_split)
        
        split_group.setLayout(split_layout)
        layout.addWidget(split_group)

        # PNG转PDF功能组
        png_group = QGroupBox("PNG 转 PDF")
        png_layout = QVBoxLayout()
        
        # 选择 PNG 输入文件夹
        self.png_input_line = QLineEdit()
        self.png_input_line.setPlaceholderText("选择 PNG 输入文件夹")
        png_layout.addWidget(self.png_input_line)
        
        self.btn_select_png_input = QPushButton("选择 PNG 输入文件夹")
        self.btn_select_png_input.clicked.connect(self.select_png_input_folder)
        png_layout.addWidget(self.btn_select_png_input)
        
        # 选择 PDF 输出文件夹
        self.png_output_line = QLineEdit()
        self.png_output_line.setPlaceholderText("选择 PDF 输出文件夹")
        png_layout.addWidget(self.png_output_line)
        
        self.btn_select_png_output = QPushButton("选择 PDF 输出文件夹")
        self.btn_select_png_output.clicked.connect(self.select_png_output_folder)
        png_layout.addWidget(self.btn_select_png_output)
        
        # 执行转换
        self.btn_convert_png = QPushButton("开始转换 PNG 到 PDF")
        self.btn_convert_png.clicked.connect(self.convert_png_to_pdf)
        png_layout.addWidget(self.btn_convert_png)
        
        png_group.setLayout(png_layout)
        layout.addWidget(png_group)

        # PDF合并功能组
        merge_group = QGroupBox("PDF 合并")
        merge_layout = QVBoxLayout()
        
        # PDF文件列表
        self.pdf_list = QListWidget()
        self.pdf_list.setMaximumHeight(150)
        merge_layout.addWidget(QLabel("选择要合并的PDF文件:"))
        merge_layout.addWidget(self.pdf_list)
        
        # 按钮布局
        merge_btn_layout = QHBoxLayout()
        self.btn_add_pdf = QPushButton("添加PDF文件夹")
        self.btn_add_pdf.clicked.connect(self.add_pdf_files)
        merge_btn_layout.addWidget(self.btn_add_pdf)
        
        self.btn_remove_pdf = QPushButton("移除选中文件")
        self.btn_remove_pdf.clicked.connect(self.remove_selected_pdf)
        merge_btn_layout.addWidget(self.btn_remove_pdf)
        
        self.btn_move_up = QPushButton("上移")
        self.btn_move_up.clicked.connect(self.move_pdf_up)
        merge_btn_layout.addWidget(self.btn_move_up)
        
        self.btn_move_down = QPushButton("下移")
        self.btn_move_down.clicked.connect(self.move_pdf_down)
        merge_btn_layout.addWidget(self.btn_move_down)
        
        merge_layout.addLayout(merge_btn_layout)
        
        # 输出文件设置
        self.merge_output_line = QLineEdit()
        self.merge_output_line.setPlaceholderText("设置合并后的PDF输出文件夹")
        merge_layout.addWidget(self.merge_output_line)
        
        merge_output_btn_layout = QHBoxLayout()
        self.btn_select_merge_output = QPushButton("选择输出文件夹")
        self.btn_select_merge_output.clicked.connect(self.select_merge_output_file)
        merge_output_btn_layout.addWidget(self.btn_select_merge_output)
        
        self.btn_merge = QPushButton("开始合并PDF")
        self.btn_merge.clicked.connect(self.merge_pdfs)
        self.btn_merge.setStyleSheet("""
            QPushButton {
                background-color: #10B981;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #059669;
            }
        """)
        merge_output_btn_layout.addWidget(self.btn_merge)
        
        merge_layout.addLayout(merge_output_btn_layout)
        
        merge_group.setLayout(merge_layout)
        layout.addWidget(merge_group)

        layout.addStretch()

    def select_pdf_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_files = [file_path]
            QMessageBox.information(self, "选择文件", f"已选择文件:\n{file_path}")

    def select_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 PDF 文件夹")
        if folder:
            self.pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            QMessageBox.information(self, "选择文件夹", f"共找到 {len(self.pdf_files)} 个 PDF 文件")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder:
            self.line_output.setText(folder)

    def select_png_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 PNG 输入文件夹")
        if folder:
            self.png_input_line.setText(folder)

    def select_png_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 PDF 输出文件夹")
        if folder:
            self.png_output_line.setText(folder)

    def split_pdf(self):
        output_folder = self.line_output.text().strip()
        if not output_folder or not os.path.exists(output_folder):
            QMessageBox.warning(self, "错误", "请先选择有效的输出文件夹")
            return
        if not self.pdf_files:
            QMessageBox.warning(self, "错误", "请先选择 PDF 文件或文件夹")
            return

        for pdf_path in self.pdf_files:
            try:
                reader = PdfReader(pdf_path)
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]

                for i, page in enumerate(reader.pages):
                    writer = PdfWriter()
                    writer.add_page(page)
                    out_path = os.path.join(output_folder, f"{base_name}_page{i+1}.pdf")
                    with open(out_path, "wb") as f:
                        writer.write(f)

            except Exception as e:
                QMessageBox.warning(self, "错误", f"处理 {pdf_path} 时出错:\n{e}")

        QMessageBox.information(self, "完成", "所有 PDF 拆分完成!")

    def convert_png_to_pdf(self):
        input_folder = self.png_input_line.text().strip()
        output_folder = self.png_output_line.text().strip()
        
        if not input_folder or not os.path.exists(input_folder):
            QMessageBox.warning(self, "错误", "请先选择有效的 PNG 输入文件夹")
            return
        if not output_folder:
            QMessageBox.warning(self, "错误", "请先选择 PDF 输出文件夹")
            return
        
        # 获取所有PNG文件
        png_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        if not png_files:
            QMessageBox.warning(self, "错误", "输入文件夹中没有找到 PNG/JPG 文件")
            return
        
        # 确保输出文件夹存在
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        converted_count = 0
        for png_file in png_files:
            try:
                # 构建完整路径
                input_path = os.path.join(input_folder, png_file)
                # 生成输出文件名（替换扩展名为.pdf）
                output_filename = os.path.splitext(png_file)[0] + '.pdf'
                output_path = os.path.join(output_folder, output_filename)
                
                # 打开图片并转换为RGB模式（确保兼容性）
                image = Image.open(input_path)
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                
                # 保存为PDF
                image.save(output_path, "PDF", resolution=100.0)
                converted_count += 1
                
            except Exception as e:
                QMessageBox.warning(self, "错误", f"转换 {png_file} 时出错:\n{e}")
        
        QMessageBox.information(self, "完成", f"转换完成! 共转换了 {converted_count} 个文件")

    def add_pdf_files(self):
        """添加PDF文件到合并列表"""
        folder = QFileDialog.getExistingDirectory(self, "选择包含PDF文件的文件夹")
        if folder:
            pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            if pdf_files:
                for file_path in pdf_files:
                    if file_path not in [self.pdf_list.item(i).text() for i in range(self.pdf_list.count())]:
                        self.pdf_list.addItem(file_path)
            else:
                QMessageBox.information(self, "提示", "文件夹中没有找到PDF文件")

    def remove_selected_pdf(self):
        """从列表中移除选中的PDF文件"""
        current_row = self.pdf_list.currentRow()
        if current_row >= 0:
            self.pdf_list.takeItem(current_row)

    def move_pdf_up(self):
        """将选中的PDF文件上移"""
        current_row = self.pdf_list.currentRow()
        if current_row > 0:
            item = self.pdf_list.takeItem(current_row)
            self.pdf_list.insertItem(current_row - 1, item)
            self.pdf_list.setCurrentRow(current_row - 1)

    def move_pdf_down(self):
        """将选中的PDF文件下移"""
        current_row = self.pdf_list.currentRow()
        if current_row < self.pdf_list.count() - 1:
            item = self.pdf_list.takeItem(current_row)
            self.pdf_list.insertItem(current_row + 1, item)
            self.pdf_list.setCurrentRow(current_row + 1)

    def select_merge_output_file(self):
        """选择合并后的PDF输出文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder:
            # 默认输出文件名为 merged_output.pdf
            output_file = os.path.join(folder, "merged_output.pdf")
            self.merge_output_line.setText(output_file)

    def merge_pdfs(self):
        """合并PDF文件"""
        output_path = self.merge_output_line.text().strip()
        if not output_path:
            QMessageBox.warning(self, "错误", "请先选择输出文件夹")
            return
        
        if self.pdf_list.count() == 0:
            QMessageBox.warning(self, "错误", "请先添加要合并的PDF文件")
            return
        
        # 获取所有PDF文件路径
        pdf_files = [self.pdf_list.item(i).text() for i in range(self.pdf_list.count())]
        
        try:
            merger = PdfWriter()
            
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        merger.add_page(page)
                else:
                    QMessageBox.warning(self, "错误", f"文件不存在: {pdf_file}")
                    return
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 写入合并后的PDF
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            
            merger.close()
            
            QMessageBox.information(self, "完成", f"PDF合并完成!\n输出文件: {output_path}")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"合并PDF时出错:\n{str(e)}")
