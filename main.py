import sys
import json
import os
import io
import re
import traceback
from copy import copy

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QLabel,
    QFileDialog, QMessageBox, QMenu, QAction, QListWidget,
    QListWidgetItem, QDialog, QFormLayout, QLineEdit, QComboBox,
    QDialogButtonBox, QAbstractItemView, QSpinBox, QSplitter,
    QGroupBox, QRadioButton, QHeaderView, QDoubleSpinBox, QButtonGroup,
    QCheckBox, QScrollArea, QSplashScreen
)
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtCore import Qt
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
from PyQt5.QtWidgets import QSplashScreen
from PyQt5.QtGui import (
    QColor, QFont, QIcon, QPixmap, QPainter
)
from PyQt5.QtCore import Qt
# 禁用 Qt 的 SSL 支持和网络探测（加速启动）
import os
os.environ['QT_QUICK_CONTROLS_STYLE'] = 'Fusion'   # 使用快速渲染引擎


def resource_path(relative_path):
    """获取资源文件的绝对路径，支持开发环境和 PyInstaller 打包后的环境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后，资源文件在临时目录中
        return os.path.join(sys._MEIPASS, relative_path)
    # 开发环境，直接使用当前目录
    return os.path.join(os.path.abspath("."), relative_path)
# ================== 全局异常钩子 ==================
def global_exception_hook(exctype, value, tb):
    error_msg = ''.join(traceback.format_exception(exctype, value, tb))
    print(error_msg)
    QMessageBox.critical(None, "程序错误", f"发生未处理异常：\n{value}")
    sys.exit(1)

sys.excepthook = global_exception_hook


# ================== 统一尺寸常量 ==================
COL_WIDTH_PX_PER_CHAR = 7.5
ROW_HEIGHT_PX_PER_PT = 1.333
DEFAULT_COL_WIDTH_CHARS = 8.0
DEFAULT_ROW_HEIGHT_PTS = 15.0
DEFAULT_COL_WIDTH_PX = int(DEFAULT_COL_WIDTH_CHARS * COL_WIDTH_PX_PER_CHAR)
DEFAULT_ROW_HEIGHT_PX = int(DEFAULT_ROW_HEIGHT_PTS * ROW_HEIGHT_PX_PER_PT)


# ================== 图片设置对话框 ==================
class ImageSetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("图片设置")
        layout = QFormLayout(self)

        self.col_width = QDoubleSpinBox()
        self.col_width.setRange(1.0, 100.0)
        self.col_width.setDecimals(1)
        self.col_width.setValue(8.0)
        layout.addRow("列宽(字符):", self.col_width)

        self.row_height = QDoubleSpinBox()
        self.row_height.setRange(1.0, 500.0)
        self.row_height.setDecimals(1)
        self.row_height.setValue(15.0)
        layout.addRow("行高(磅):", self.row_height)

        self.rotation = QDoubleSpinBox()
        self.rotation.setRange(-360.0, 360.0)
        self.rotation.setDecimals(1)
        self.rotation.setValue(0.0)
        layout.addRow("旋转角度(正顺时针):", self.rotation)

        self.width_scale = QDoubleSpinBox()
        self.width_scale.setRange(0.1, 1.0)
        self.width_scale.setSingleStep(0.1)
        self.width_scale.setValue(1.0)
        layout.addRow("宽度缩放比例:", self.width_scale)

        self.height_scale = QDoubleSpinBox()
        self.height_scale.setRange(0.1, 1.0)
        self.height_scale.setSingleStep(0.1)
        self.height_scale.setValue(1.0)
        layout.addRow("高度缩放比例:", self.height_scale)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_values(self):
        return (self.col_width.value(), self.row_height.value(),
                self.rotation.value(),
                self.width_scale.value(), self.height_scale.value())


# ================== 数据转换规则对话框 ==================
class TransformDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("数据转换规则")
        layout = QVBoxLayout(self)

        self.btn_group = QButtonGroup(self)
        self.radio_none = QRadioButton("无转换")
        self.radio_div1000 = QRadioButton("除以1000 (ms→s)")
        self.radio_strip = QRadioButton("去除末尾字母")
        self.radio_custom = QRadioButton("自定义表达式")
        self.btn_group.addButton(self.radio_none, 0)
        self.btn_group.addButton(self.radio_div1000, 1)
        self.btn_group.addButton(self.radio_strip, 2)
        self.btn_group.addButton(self.radio_custom, 3)
        self.radio_none.setChecked(True)

        layout.addWidget(self.radio_none)
        layout.addWidget(self.radio_div1000)
        layout.addWidget(self.radio_strip)
        layout.addWidget(self.radio_custom)

        self.custom_expr = QLineEdit()
        self.custom_expr.setPlaceholderText("输入Python表达式，如 x/1000")
        self.custom_expr.setEnabled(False)
        self.radio_custom.toggled.connect(lambda checked: self.custom_expr.setEnabled(checked))
        layout.addWidget(self.custom_expr)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_transform(self):
        if self.radio_none.isChecked():
            return 'none', ''
        elif self.radio_div1000.isChecked():
            return 'div1000', ''
        elif self.radio_strip.isChecked():
            return 'strip_letters', ''
        else:
            return 'custom', self.custom_expr.text()


# ================== 归档配置对话框（动态输入框） ==================
class ArchiveConfigDialog(QDialog):
    def __init__(self, source_wb, template_range=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("归档配置")
        self.setMinimumSize(800, 650)
        self.source_wb = source_wb
        self.selected_range = None
        self.header_inputs = []  # 动态生成的表头输入框

        main_layout = QVBoxLayout(self)
        # 模板区域信息
        if template_range:
            min_row, min_col, max_row, max_col = template_range
            rows = max_row - min_row + 1
            info_text = f"📌 已选首列范围：{rows} 行（{get_column_letter(min_col)}{min_row}:{get_column_letter(min_col)}{max_row}）"
        else:
            info_text = "⚠️ 未获取到模板区域，请重新框选"
        info_label = QLabel(info_text)
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; background-color: #ecf0f1; padding: 5px;")
        main_layout.addWidget(info_label)

        # 表头行数选择（改变时动态生成输入框）
        form = QFormLayout()
        self.header_rows_spin = QSpinBox()
        self.header_rows_spin.setMinimum(1)
        self.header_rows_spin.setValue(1)
        self.header_rows_spin.valueChanged.connect(self.on_header_rows_changed)
        form.addRow("表头行数:", self.header_rows_spin)
        main_layout.addLayout(form)

        # 动态输入框容器
        self.header_group = QGroupBox("新表头内容（每行一个输入框）")
        self.header_layout = QVBoxLayout(self.header_group)
        self.header_scroll = QScrollArea()
        self.header_scroll.setWidgetResizable(True)
        self.header_scroll.setWidget(self.header_group)
        main_layout.addWidget(self.header_scroll)

        # 源数据选择
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("源数据Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(source_wb.sheetnames)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(self.sheet_combo)
        sheet_layout.addStretch()
        main_layout.addLayout(sheet_layout)

        self.table = QTableWidget()
        self.table.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        main_layout.addWidget(self.table)

        bottom_layout = QHBoxLayout()
        self.range_label = QLabel("未选择区域")
        bottom_layout.addWidget(self.range_label)
        self.hint_label = QLabel("")
        bottom_layout.addWidget(self.hint_label)
        bottom_layout.addStretch()
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        bottom_layout.addWidget(buttons)
        main_layout.addLayout(bottom_layout)

        self.load_sheet(source_wb.sheetnames[0])
        self.on_header_rows_changed(1)  # 初始化输入框

    def on_header_rows_changed(self, value):
        # 清空并重新生成输入框
        for i in reversed(range(self.header_layout.count())):
            widget = self.header_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.header_inputs.clear()
        for i in range(value):
            edit = QLineEdit()
            edit.setPlaceholderText(f"第{i+1}行表头")
            self.header_layout.addWidget(edit)
            self.header_inputs.append(edit)

    def on_sheet_changed(self, index):
        self.load_sheet(self.sheet_combo.currentText())

    def load_sheet(self, sheet_name):
        ws = self.source_wb[sheet_name]
        self.table.clear()
        self.table.clearSpans()
        real_max_row = ws.max_row
        real_max_col = ws.max_column
        if real_max_row > 200:
            for r in range(real_max_row, 0, -1):
                if any(cell.value is not None for cell in ws[r]):
                    real_max_row = r
                    break
            real_max_row += 2
        if real_max_col > 30:
            for c in range(real_max_col, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for r in range(1, real_max_row+1)):
                    real_max_col = c
                    break
            real_max_col += 2
        max_rows = min(real_max_row, 500)
        max_cols = min(real_max_col, 100)

        self.table.setRowCount(max_rows)
        self.table.setColumnCount(max_cols)
        self.table.setSelectionMode(QAbstractItemView.ContiguousSelection)

        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            if min_row > max_rows or min_col > max_cols:
                continue
            span_rows = min(max_row, max_rows) - min_row + 1
            span_cols = min(max_col, max_cols) - min_col + 1
            self.table.setSpan(min_row-1, min_col-1, span_rows, span_cols)

        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols), start=1):
            for col_idx, cell in enumerate(row, start=1):
                if cell.value is not None:
                    item = QTableWidgetItem(str(cell.value))
                    self.table.setItem(row_idx-1, col_idx-1, item)

        self._apply_uniform_sizes(ws, max_cols, max_rows)
        self.selected_range = None
        self.range_label.setText("未选择区域")
        self.hint_label.setText("")

    def _apply_uniform_sizes(self, ws, max_cols, max_rows):
        for col in range(1, max_cols + 1):
            letter = get_column_letter(col)
            if letter in ws.column_dimensions and ws.column_dimensions[letter].width:
                width = int(ws.column_dimensions[letter].width * COL_WIDTH_PX_PER_CHAR)
            else:
                width = DEFAULT_COL_WIDTH_PX
            self.table.setColumnWidth(col - 1, width)
        for row in range(1, max_rows + 1):
            if row in ws.row_dimensions and ws.row_dimensions[row].height:
                height = int(ws.row_dimensions[row].height * ROW_HEIGHT_PX_PER_PT)
            else:
                height = DEFAULT_ROW_HEIGHT_PX
            self.table.setRowHeight(row - 1, height)

    def on_selection_changed(self):
        indexes = self.table.selectedIndexes()
        if not indexes:
            self.selected_range = None
            self.range_label.setText("未选择区域")
            return
        top = min(idx.row() for idx in indexes) + 1
        left = min(idx.column() for idx in indexes) + 1
        bottom = max(idx.row() for idx in indexes) + 1
        right = max(idx.column() for idx in indexes) + 1
        self.selected_range = (top, left, bottom, right)
        rows = bottom - top + 1
        cols = right - left + 1
        addr = f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}  ({rows}行×{cols}列)"
        self.range_label.setText(addr)

    def get_selection(self):
        header_rows = self.header_rows_spin.value()
        headers = [edit.text() for edit in self.header_inputs]
        sheet = self.sheet_combo.currentText()
        if self.selected_range:
            min_row, min_col, max_row, max_col = self.selected_range
            range_str = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        else:
            range_str = ""
        return header_rows, headers, sheet, range_str


# ================== JMP配置对话框（动态列数） ==================
class JMPConfigDialog(QDialog):
    def __init__(self, template_wb, anchor_cell, parent=None):
        super().__init__(parent)
        self.setWindowTitle("JMP数据区配置")
        self.setMinimumSize(600, 450)
        self.template_wb = template_wb
        self.anchor_cell = anchor_cell
        self.header_inputs = []  # 动态生成的表头输入框

        layout = QVBoxLayout(self)

        info_label = QLabel(f"📍 锚点单元格: {anchor_cell}")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        layout.addWidget(info_label)

        # 表头列数选择
        header_form = QFormLayout()
        self.header_cols_spin = QSpinBox()
        self.header_cols_spin.setMinimum(1)
        self.header_cols_spin.setValue(2)
        self.header_cols_spin.valueChanged.connect(self.on_header_cols_changed)
        header_form.addRow("表头列数:", self.header_cols_spin)
        layout.addLayout(header_form)

        # 动态输入框容器
        self.header_group = QGroupBox("表头内容（每列一个输入框）")
        self.header_layout = QVBoxLayout(self.header_group)
        self.header_scroll = QScrollArea()
        self.header_scroll.setWidgetResizable(True)
        self.header_scroll.setWidget(self.header_group)
        layout.addWidget(self.header_scroll)

        # 源数据区域
        source_group = QGroupBox("源数据区域（从模板Sheet中选择）")
        source_layout = QFormLayout(source_group)
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(template_wb.sheetnames)
        source_layout.addRow("源Sheet:", self.sheet_combo)

        range_layout = QHBoxLayout()
        self.range_edit = QLineEdit()
        self.range_edit.setPlaceholderText("例如 C2:E10，或点击右侧按钮框选")
        range_layout.addWidget(self.range_edit)
        btn_select = QPushButton("框选区域")
        btn_select.clicked.connect(self.select_range)
        range_layout.addWidget(btn_select)
        source_layout.addRow("源区域:", range_layout)
        layout.addWidget(source_group)

        self.merge_cols_check = QCheckBox("将多列数据拼接为单列（先行后列）")
        layout.addWidget(self.merge_cols_check)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.on_header_cols_changed(2)  # 初始化输入框

    def on_header_cols_changed(self, value):
        # 清空并重新生成输入框
        for i in reversed(range(self.header_layout.count())):
            widget = self.header_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.header_inputs.clear()
        for i in range(value):
            edit = QLineEdit()
            edit.setPlaceholderText(f"第{i+1}列表头")
            self.header_layout.addWidget(edit)
            self.header_inputs.append(edit)

    def select_range(self):
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name:
            return
        dlg = SourceSelectDialog(self.template_wb, self)
        dlg.sheet_combo.setCurrentText(sheet_name)
        if dlg.exec_():
            sel_sheet, sel_range = dlg.get_selection()
            if sel_range:
                self.range_edit.setText(sel_range)

    def get_config(self):
        headers = [edit.text() for edit in self.header_inputs]
        return {
            'header_cols': headers,
            'source_sheet': self.sheet_combo.currentText(),
            'source_range': self.range_edit.text(),
            'merge_columns': self.merge_cols_check.isChecked()
        }


# ================== 源数据选择对话框 ==================
class SourceSelectDialog(QDialog):
    def __init__(self, source_wb, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择数据源区域")
        self.setMinimumSize(800, 600)
        self.source_wb = source_wb
        self.selected_range = None

        main_layout = QVBoxLayout(self)
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("源Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(source_wb.sheetnames)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        top_layout.addWidget(self.sheet_combo)
        top_layout.addStretch()
        main_layout.addLayout(top_layout)

        self.table = QTableWidget()
        self.table.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        main_layout.addWidget(self.table)

        bottom_layout = QHBoxLayout()
        self.range_label = QLabel("未选择区域")
        bottom_layout.addWidget(self.range_label)
        bottom_layout.addStretch()
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        bottom_layout.addWidget(buttons)
        main_layout.addLayout(bottom_layout)

        self.load_sheet(source_wb.sheetnames[0])

    def on_sheet_changed(self, index):
        self.load_sheet(self.sheet_combo.currentText())

    def load_sheet(self, sheet_name):
        ws = self.source_wb[sheet_name]
        self.table.clear()
        self.table.clearSpans()
        real_max_row = ws.max_row
        real_max_col = ws.max_column
        if real_max_row > 200:
            for r in range(real_max_row, 0, -1):
                if any(cell.value is not None for cell in ws[r]):
                    real_max_row = r
                    break
            real_max_row += 2
        if real_max_col > 30:
            for c in range(real_max_col, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for r in range(1, real_max_row+1)):
                    real_max_col = c
                    break
            real_max_col += 2
        max_rows = min(real_max_row, 500)
        max_cols = min(real_max_col, 100)

        self.table.setRowCount(max_rows)
        self.table.setColumnCount(max_cols)

        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            if min_row > max_rows or min_col > max_cols:
                continue
            span_rows = min(max_row, max_rows) - min_row + 1
            span_cols = min(max_col, max_cols) - min_col + 1
            self.table.setSpan(min_row-1, min_col-1, span_rows, span_cols)

        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols), start=1):
            for col_idx, cell in enumerate(row, start=1):
                if cell.value is not None:
                    item = QTableWidgetItem(str(cell.value))
                    self.table.setItem(row_idx-1, col_idx-1, item)

        self._apply_uniform_sizes(ws, max_cols, max_rows)
        self.selected_range = None
        self.range_label.setText("未选择区域")

    def _apply_uniform_sizes(self, ws, max_cols, max_rows):
        for col in range(1, max_cols + 1):
            letter = get_column_letter(col)
            if letter in ws.column_dimensions and ws.column_dimensions[letter].width:
                width = int(ws.column_dimensions[letter].width * COL_WIDTH_PX_PER_CHAR)
            else:
                width = DEFAULT_COL_WIDTH_PX
            self.table.setColumnWidth(col - 1, width)
        for row in range(1, max_rows + 1):
            if row in ws.row_dimensions and ws.row_dimensions[row].height:
                height = int(ws.row_dimensions[row].height * ROW_HEIGHT_PX_PER_PT)
            else:
                height = DEFAULT_ROW_HEIGHT_PX
            self.table.setRowHeight(row - 1, height)

    def on_selection_changed(self):
        indexes = self.table.selectedIndexes()
        if not indexes:
            self.selected_range = None
            self.range_label.setText("未选择区域")
            return
        top = min(idx.row() for idx in indexes) + 1
        left = min(idx.column() for idx in indexes) + 1
        bottom = max(idx.row() for idx in indexes) + 1
        right = max(idx.column() for idx in indexes) + 1
        self.selected_range = (top, left, bottom, right)
        addr = f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}"
        self.range_label.setText(addr)

    def get_selection(self):
        sheet_name = self.sheet_combo.currentText()
        if self.selected_range:
            min_row, min_col, max_row, max_col = self.selected_range
            range_str = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        else:
            range_str = ""
        return sheet_name, range_str


# ================== 内部图片选择对话框 ==================
class InternalImageSelectDialog(QDialog):
    def __init__(self, cached_images, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择图片（可多选）")
        self.setMinimumSize(650, 450)
        self.cached_images = cached_images

        layout = QVBoxLayout(self)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("筛选Sheet:"))
        self.sheet_filter = QComboBox()
        sheets = sorted(list(set(img[0] for img in cached_images)))
        self.sheet_filter.addItem("全部")
        self.sheet_filter.addItems(sheets)
        self.sheet_filter.currentIndexChanged.connect(self.filter_images)
        filter_layout.addWidget(self.sheet_filter)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["位置", "Sheet", "尺寸"])
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.populate_table(cached_images)

    def populate_table(self, images):
        self.table.setRowCount(len(images))
        for i, (sheet, idx, pos, data, w, h) in enumerate(images):
            self.table.setItem(i, 0, QTableWidgetItem(pos))
            self.table.setItem(i, 1, QTableWidgetItem(sheet))
            self.table.setItem(i, 2, QTableWidgetItem(f"{w}×{h}"))
        self.table.resizeColumnsToContents()

    def filter_images(self):
        filter_sheet = self.sheet_filter.currentText()
        if filter_sheet == "全部":
            filtered = self.cached_images
        else:
            filtered = [img for img in self.cached_images if img[0] == filter_sheet]
        self.populate_table(filtered)

    def get_selected_images(self):
        selected_rows = set(idx.row() for idx in self.table.selectedIndexes())
        filter_sheet = self.sheet_filter.currentText()
        if filter_sheet == "全部":
            current_images = self.cached_images
        else:
            current_images = [img for img in self.cached_images if img[0] == filter_sheet]
        result = []
        for row in sorted(selected_rows):
            if row < len(current_images):
                result.append(current_images[row])
        return result


# ================== 批量图片对话框 ==================
class BatchImageDialog(QDialog):
    def __init__(self, cached_images, rows, cols, parent=None):
        super().__init__(parent)
        self.setWindowTitle("批量图片配置")
        self.setMinimumSize(750, 500)
        self.cached_images = cached_images
        self.target_rows = rows
        self.target_cols = cols
        self.image_list = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"目标区域：{rows} 行 × {cols} 列，共需 {rows*cols} 张图片"))

        source_group = QGroupBox("图片来源")
        source_layout = QVBoxLayout(source_group)
        self.radio_internal = QRadioButton("从数据源提取")
        self.radio_file = QRadioButton("从外部文件选择")
        self.radio_internal.setChecked(True)
        source_layout.addWidget(self.radio_internal)
        source_layout.addWidget(self.radio_file)
        layout.addWidget(source_group)

        self.image_table = QTableWidget()
        self.image_table.setColumnCount(2)
        self.image_table.setHorizontalHeaderLabels(["序号", "来源"])
        self.image_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.image_table)

        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("添加图片")
        self.btn_add.clicked.connect(self.add_images)
        btn_layout.addWidget(self.btn_add)
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.clicked.connect(self.clear_images)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def add_images(self):
        if self.radio_internal.isChecked():
            dlg = InternalImageSelectDialog(self.cached_images, self)
            if dlg.exec_():
                selected = dlg.get_selected_images()
                for item in selected:
                    self.image_list.append(('internal', item))
                    row = self.image_table.rowCount()
                    self.image_table.insertRow(row)
                    self.image_table.setItem(row, 0, QTableWidgetItem(str(row+1)))
                    self.image_table.setItem(row, 1, QTableWidgetItem(item[2]))
        else:
            files, _ = QFileDialog.getOpenFileNames(self, "选择图片文件", "",
                                                    "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif);;所有文件 (*)")
            for f in files:
                self.image_list.append(('file', f))
                row = self.image_table.rowCount()
                self.image_table.insertRow(row)
                self.image_table.setItem(row, 0, QTableWidgetItem(str(row+1)))
                self.image_table.setItem(row, 1, QTableWidgetItem(os.path.basename(f)))

    def clear_images(self):
        self.image_list.clear()
        self.image_table.setRowCount(0)

    def get_image_sequence(self):
        return self.image_list


# ================== 主窗口 ==================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动报告工具")
        self.setGeometry(100, 100, 1400, 850)

        self.template_wb = None
        self.source_wb = None
        self.template_path = None
        self.source_path = None
        self.current_sheet_name = None
        self.mappings = []
        self.current_selection = None
        self._image_streams = []
        self.cached_images = []

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # Sheet面板
        sheet_panel = QWidget()
        sheet_layout = QVBoxLayout(sheet_panel)
        sheet_layout.setContentsMargins(5, 5, 5, 5)
        sheet_label = QLabel("模板Sheet列表")
        sheet_label.setStyleSheet("font-weight: bold;")
        sheet_layout.addWidget(sheet_label)
        self.sheet_list = QListWidget()
        self.sheet_list.itemClicked.connect(self.on_sheet_clicked)
        sheet_layout.addWidget(self.sheet_list)

        # 中间
        center_widget = QWidget()
        center_layout = QVBoxLayout(center_widget)
        btn_layout = QHBoxLayout()
        self.btn_open_template = QPushButton("打开模板文件")
        self.btn_open_template.clicked.connect(self.open_template)
        btn_layout.addWidget(self.btn_open_template)
        self.btn_open_source = QPushButton("打开数据源文件")
        self.btn_open_source.clicked.connect(self.open_source)
        btn_layout.addWidget(self.btn_open_source)
        center_layout.addLayout(btn_layout)
        self.table = QTableWidget()
        self.table.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        center_layout.addWidget(self.table)

        # 右侧
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("当前Sheet映射列表"))
        self.mapping_list = QListWidget()
        right_layout.addWidget(self.mapping_list)

        self.btn_confirm_mapping = QPushButton("确认映射")
        self.btn_confirm_mapping.clicked.connect(self.confirm_mapping)
        right_layout.addWidget(self.btn_confirm_mapping)

        self.btn_output_report = QPushButton("输出报告")
        self.btn_output_report.clicked.connect(self.output_report)
        right_layout.addWidget(self.btn_output_report)

        self.btn_save_config = QPushButton("保存配置为模板")
        self.btn_save_config.clicked.connect(self.save_config)
        right_layout.addWidget(self.btn_save_config)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(sheet_panel)
        splitter.addWidget(center_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([220, 700, 280])
        main_layout.addWidget(splitter)

    # ==================== 文件加载 ====================
    def open_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择模板", "", "Excel文件 (*.xlsx)")
        if not path:
            return
        self.template_path = path
        try:
            self.template_wb = openpyxl.load_workbook(path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开模板文件：{e}")
            return

        config_path = path.rsplit('.', 1)[0] + '_config.json'
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.mappings = config.get('mappings', [])
            except Exception as e:
                QMessageBox.warning(self, "警告", f"配置文件读取失败：{e}")
                self.mappings = []
        else:
            self.mappings = []

        self.sheet_list.clear()
        for name in self.template_wb.sheetnames:
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.sheet_list.addItem(item)

        if self.sheet_list.count() > 0:
            self.sheet_list.setCurrentRow(0)
            self.current_sheet_name = self.template_wb.sheetnames[0]
            self.display_sheet(self.template_wb[self.current_sheet_name])

    def open_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择数据源", "", "Excel文件 (*.xlsx)")
        if not path:
            return
        self.source_path = path
        try:
            self.source_wb = openpyxl.load_workbook(path)
            self._cache_all_images()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开数据源文件：{e}")
            return
        QMessageBox.information(self, "提示", f"数据源文件已加载：{os.path.basename(path)}")

    def _cache_all_images(self):
        self.cached_images = []
        for sheet_name in self.source_wb.sheetnames:
            ws = self.source_wb[sheet_name]
            for idx, img in enumerate(ws._images):
                if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                    from_cell = img.anchor._from
                    row = from_cell.row + 1
                    col = from_cell.col + 1
                    pos = f"{get_column_letter(col)}{row}"
                else:
                    pos = "未知"
                try:
                    img_data = img._data()
                except AttributeError:
                    try:
                        img_data = img.ref.read()
                    except Exception as e:
                        print(f"警告：无法读取图片 {pos}：{e}")
                        continue
                width = getattr(img, 'width', '')
                height = getattr(img, 'height', '')
                self.cached_images.append((sheet_name, idx, pos, img_data, width, height))

    # ==================== Sheet渲染 ====================
    def on_sheet_clicked(self, item):
        self.current_sheet_name = item.text()
        self.display_sheet(self.template_wb[self.current_sheet_name])
        self.refresh_mapping_list()

    def display_sheet(self, ws):
        self.table.clear()
        self.table.clearSpans()
        real_max_row = ws.max_row
        real_max_col = ws.max_column
        if real_max_row > 200:
            for r in range(real_max_row, 0, -1):
                if any(cell.value is not None for cell in ws[r]):
                    real_max_row = r
                    break
            real_max_row += 2
        if real_max_col > 30:
            for c in range(real_max_col, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for r in range(1, real_max_row+1)):
                    real_max_col = c
                    break
            real_max_col += 2

        max_rows = min(real_max_row, 500)
        max_cols = min(real_max_col, 100)
        self.table.setRowCount(max_rows)
        self.table.setColumnCount(max_cols)

        try:
            for merged_range in ws.merged_cells.ranges:
                min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                if min_row > max_rows or min_col > max_cols:
                    continue
                span_rows = min(max_row, max_rows) - min_row + 1
                span_cols = min(max_col, max_cols) - min_col + 1
                self.table.setSpan(min_row-1, min_col-1, span_rows, span_cols)

            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    if cell.value is None:
                        continue
                    item = QTableWidgetItem(str(cell.value))
                    if cell.font:
                        font = QFont()
                        font.setFamily(cell.font.name or 'Arial')
                        size = cell.font.size or 10
                        if size is not None:
                            font.setPointSize(int(size))
                        font.setBold(cell.font.bold)
                        item.setFont(font)
                    if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                        raw = cell.fill.fgColor.rgb
                        try:
                            rgb = str(raw) if not isinstance(raw, str) else raw
                            if len(rgb) == 8 and rgb.startswith('00'):
                                rgb = rgb[2:]
                            color = QColor(f"#{rgb}")
                            if color.isValid():
                                item.setBackground(color)
                        except:
                            pass
                    self.table.setItem(row_idx-1, col_idx-1, item)
        except Exception as e:
            QMessageBox.warning(self, "渲染警告", f"表格渲染出现异常：{e}")

        self._apply_uniform_sizes(ws, max_cols, max_rows)

    def _apply_uniform_sizes(self, ws, max_cols, max_rows):
        for col in range(1, max_cols + 1):
            letter = get_column_letter(col)
            if letter in ws.column_dimensions and ws.column_dimensions[letter].width:
                width = int(ws.column_dimensions[letter].width * COL_WIDTH_PX_PER_CHAR)
            else:
                width = DEFAULT_COL_WIDTH_PX
            self.table.setColumnWidth(col - 1, width)
        for row in range(1, max_rows + 1):
            if row in ws.row_dimensions and ws.row_dimensions[row].height:
                height = int(ws.row_dimensions[row].height * ROW_HEIGHT_PX_PER_PT)
            else:
                height = DEFAULT_ROW_HEIGHT_PX
            self.table.setRowHeight(row - 1, height)

    # ==================== 右键菜单 ====================
    def show_context_menu(self, pos):
        indexes = self.table.selectedIndexes()
        if not indexes:
            return
        top = min(idx.row() for idx in indexes) + 1
        left = min(idx.column() for idx in indexes) + 1
        bottom = max(idx.row() for idx in indexes) + 1
        right = max(idx.column() for idx in indexes) + 1
        self.current_selection = (top, left, bottom, right)

        menu = QMenu()
        info_action = QAction(
            f"选中区域: {get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}  ({bottom-top+1}行×{right-left+1}列)",
            self
        )
        info_action.setEnabled(False)
        menu.addAction(info_action)
        menu.addSeparator()
        menu.addAction(QAction("设为数据填充区", self, triggered=self.add_data_mapping))
        menu.addAction(QAction("设为图片区域", self, triggered=self.add_image_mapping))
        menu.addAction(QAction("设为归档区域（右移）", self, triggered=self.add_archive_mapping))
        menu.addAction(QAction("设为JMP数据区", self, triggered=self.add_jmp_mapping))
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    # ==================== 映射添加 ====================
    def add_data_mapping(self):
        if not self.current_selection or not self.source_wb:
            QMessageBox.warning(self, "提示", "请先打开数据源文件并选中目标区域")
            return
        dlg = SourceSelectDialog(self.source_wb, self)
        if dlg.exec_():
            src_sheet, src_range = dlg.get_selection()
            if not src_range:
                QMessageBox.warning(self, "提示", "源区域不能为空")
                return

            trans_dlg = TransformDialog(self)
            if not trans_dlg.exec_():
                return
            trans_type, trans_expr = trans_dlg.get_transform()

            self.mappings.append({
                'type': 'data',
                'target_sheet': self.current_sheet_name,
                'target_range': self.current_selection,
                'source_sheet': src_sheet,
                'source_range': src_range,
                'transform': trans_type,
                'transform_expr': trans_expr
            })
            self.refresh_mapping_list()

    def add_image_mapping(self):
        if not self.current_selection:
            QMessageBox.warning(self, "提示", "请先在模板中选中目标区域")
            return

        dlg = ImageSetupDialog(self)
        if not dlg.exec_():
            return
        col_width_chars, row_height_pts, rotation, w_scale, h_scale = dlg.get_values()

        t_min_row, t_min_col, t_max_row, t_max_col = self.current_selection
        rows = t_max_row - t_min_row + 1
        cols = t_max_col - t_min_col + 1
        if rows * cols == 1:
            self._add_single_image_mapping(col_width_chars, row_height_pts, rotation, w_scale, h_scale)
        else:
            self._add_batch_image_mapping(rows, cols, col_width_chars, row_height_pts, rotation, w_scale, h_scale)

    def _add_single_image_mapping(self, col_width_chars, row_height_pts, rotation, w_scale, h_scale):
        t_row, t_col, _, _ = self.current_selection
        anchor = f"{get_column_letter(t_col)}{t_row}"

        if not self.source_wb:
            QMessageBox.information(self, "提示", "未打开数据源文件，只能从外部文件选择图片")
            img_path, _ = QFileDialog.getOpenFileName(self, "选择图片文件", "",
                                                       "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif);;所有文件 (*)")
            if img_path:
                mapping = {
                    'type': 'image',
                    'target_sheet': self.current_sheet_name,
                    'anchor_cell': anchor,
                    'image_path': img_path,
                    'col_width_chars': col_width_chars,
                    'row_height_pts': row_height_pts,
                    'rotation': rotation,
                    'width_scale': w_scale,
                    'height_scale': h_scale
                }
                self.mappings.append(mapping)
                self.refresh_mapping_list()
            return

        if not self.cached_images:
            reply = QMessageBox.question(self, "图片来源", "数据源中没有图片，是否从外部文件选择？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                img_path, _ = QFileDialog.getOpenFileName(self, "选择图片文件", "",
                                                           "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif);;所有文件 (*)")
                if img_path:
                    mapping = {
                        'type': 'image',
                        'target_sheet': self.current_sheet_name,
                        'anchor_cell': anchor,
                        'image_path': img_path,
                        'col_width_chars': col_width_chars,
                        'row_height_pts': row_height_pts,
                        'rotation': rotation,
                        'width_scale': w_scale,
                        'height_scale': h_scale
                    }
                    self.mappings.append(mapping)
                    self.refresh_mapping_list()
            return

        dlg = InternalImageSelectDialog(self.cached_images, self)
        if dlg.exec_():
            selected = dlg.get_selected_images()
            if selected:
                img_info = selected[0]
                mapping = {
                    'type': 'image',
                    'target_sheet': self.current_sheet_name,
                    'anchor_cell': anchor,
                    'image_bytes': img_info[3],
                    'orig_width': img_info[4],
                    'orig_height': img_info[5],
                    'col_width_chars': col_width_chars,
                    'row_height_pts': row_height_pts,
                    'rotation': rotation,
                    'width_scale': w_scale,
                    'height_scale': h_scale
                }
                self.mappings.append(mapping)
                self.refresh_mapping_list()

    def _add_batch_image_mapping(self, rows, cols, col_width_chars, row_height_pts, rotation, w_scale, h_scale):
        if not self.source_wb:
            QMessageBox.warning(self, "提示", "请先打开数据源文件")
            return
        if not self.cached_images:
            QMessageBox.warning(self, "提示", "数据源中没有图片，无法批量添加")
            return
        dlg = BatchImageDialog(self.cached_images, rows, cols, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        images = dlg.get_image_sequence()
        if len(images) < rows * cols:
            QMessageBox.warning(self, "图片不足", f"需要 {rows*cols} 张图片，只提供了 {len(images)} 张")
            return
        t_min_row, t_min_col, _, _ = self.current_selection
        idx = 0
        for r in range(rows):
            for c in range(cols):
                anchor = f"{get_column_letter(t_min_col + c)}{t_min_row + r}"
                img_type, img_data = images[idx]
                idx += 1
                base_mapping = {
                    'type': 'image',
                    'target_sheet': self.current_sheet_name,
                    'anchor_cell': anchor,
                    'col_width_chars': col_width_chars,
                    'row_height_pts': row_height_pts,
                    'rotation': rotation,
                    'width_scale': w_scale,
                    'height_scale': h_scale
                }
                if img_type == 'file':
                    base_mapping['image_path'] = img_data
                else:
                    base_mapping['image_bytes'] = img_data[3]
                    base_mapping['orig_width'] = img_data[4]
                    base_mapping['orig_height'] = img_data[5]
                self.mappings.append(base_mapping)
        self.refresh_mapping_list()

    def add_archive_mapping(self):
        if not self.current_selection or not self.source_wb:
            QMessageBox.warning(self, "提示", "请先打开数据源文件并选中归档区域的首列")
            return
        t_min_row, t_min_col, t_max_row, _ = self.current_selection
        ws = self.template_wb[self.current_sheet_name]

        dlg = ArchiveConfigDialog(self.source_wb, template_range=(t_min_row, t_min_col, t_max_row, t_min_col), parent=self)
        if dlg.exec_():
            header_rows, headers, src_sheet, src_range = dlg.get_selection()
            if not src_range:
                QMessageBox.warning(self, "提示", "源数据区域不能为空")
                return

            max_col = ws.max_column
            right_col = t_min_col
            for col in range(t_min_col, min(t_min_col + 50, max_col + 1)):
                has_value = False
                for r in range(t_min_row, t_min_row + header_rows):
                    cell = ws.cell(row=r, column=col)
                    if cell.value is not None:
                        has_value = True
                        break
                if has_value:
                    right_col = col
                else:
                    break

            full_block = (t_min_row, t_min_col, t_max_row, right_col)
            block_str = f"{get_column_letter(t_min_col)}{t_min_row}:{get_column_letter(right_col)}{t_max_row}"
            reply = QMessageBox.question(self, "确认归档范围",
                f"自动检测到归档区域：\n{block_str}\n（首列 {get_column_letter(t_min_col)} → 末列 {get_column_letter(right_col)}）\n\n是否继续？",
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

            block_rows = t_max_row - t_min_row + 1
            data_rows = block_rows - header_rows
            s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(src_range)
            src_rows = s_max_row - s_min_row + 1
            if src_rows != data_rows:
                QMessageBox.warning(self, "行数不匹配",
                    f"源数据区域需选择 {data_rows} 行，当前选择了 {src_rows} 行")
                return
            # 表头内容已经由输入框保证数量一致，无需检查

            self.mappings.append({
                'type': 'archive_shift_right',
                'target_sheet': self.current_sheet_name,
                'block_range': full_block,
                'header_rows': header_rows,
                'new_headers': headers,
                'source_sheet': src_sheet,
                'source_range': src_range
            })
            self.refresh_mapping_list()

    def add_jmp_mapping(self):
        if not self.current_selection:
            QMessageBox.warning(self, "提示", "请先选中JMP数据区的起始单元格（锚点）")
            return
        t_row, t_col, _, _ = self.current_selection
        anchor = f"{get_column_letter(t_col)}{t_row}"

        dlg = JMPConfigDialog(self.template_wb, anchor, self)
        if dlg.exec_():
            config = dlg.get_config()
            if not config['source_range']:
                QMessageBox.warning(self, "提示", "源区域不能为空")
                return

            mapping = {
                'type': 'jmp',
                'target_sheet': self.current_sheet_name,
                'anchor_cell': anchor,
                'header_cols': config['header_cols'],  # 列表
                'source_sheet': config['source_sheet'],
                'source_range': config['source_range'],
                'merge_columns': config['merge_columns']
            }
            self.mappings.append(mapping)
            self.refresh_mapping_list()

    # ==================== 映射列表刷新 ====================
    def refresh_mapping_list(self):
        self.mapping_list.clear()
        for i, m in enumerate(self.mappings):
            if m.get('target_sheet') != self.current_sheet_name:
                continue
            desc = f"{i+1}. "
            if m['type'] == 'data':
                trans = m.get('transform', 'none')
                trans_str = f" [转换:{trans}]" if trans != 'none' else ""
                desc += f"数据: {m['target_range']} <- {m['source_sheet']}!{m['source_range']}{trans_str}"
            elif m['type'] == 'image':
                desc += f"图片: 锚点{m['anchor_cell']} (列宽{m.get('col_width_chars','?')} 行高{m.get('row_height_pts','?')} 旋转{m.get('rotation',0)}° 缩放{m.get('width_scale',1.0)}x{m.get('height_scale',1.0)})"
            elif m['type'] == 'archive_shift_right':
                desc += f"归档: {m['block_range']} 新表头\"{m['new_headers']}\""
            elif m['type'] == 'jmp':
                headers_str = ','.join(m['header_cols'])
                desc += f"JMP: 锚点{m['anchor_cell']} <- {m['source_sheet']}!{m['source_range']} (表头:{headers_str} 拼接:{m['merge_columns']})"
            self.mapping_list.addItem(desc)

    # ==================== 确认映射 ====================
    def confirm_mapping(self):
        if not self.mappings:
            QMessageBox.information(self, "提示", "当前没有任何映射，请先添加映射")
            return
        current_mappings = [m for m in self.mappings if m.get('target_sheet') == self.current_sheet_name]
        if not current_mappings:
            QMessageBox.information(self, "提示", f"当前Sheet“{self.current_sheet_name}”没有映射")
        else:
            QMessageBox.information(self, "确认", f"已确认当前Sheet“{self.current_sheet_name}”的 {len(current_mappings)} 条映射，待输出报告时统一执行。")

    # ==================== 输出报告 ====================
    def output_report(self):
        if not self.template_wb:
            return
        checked_sheets = [self.sheet_list.item(i).text() for i in range(self.sheet_list.count())
                          if self.sheet_list.item(i).checkState() == Qt.Checked]
        if not checked_sheets:
            QMessageBox.warning(self, "警告", "请至少勾选一个需要输出的Sheet")
            return

        data_template_wb = openpyxl.load_workbook(self.template_path, data_only=True)
        data_source_wb = openpyxl.load_workbook(self.source_path, data_only=True) if self.source_path else None

        failed_mappings = []
        for sheet_name in checked_sheets:
            ws = self.template_wb[sheet_name]
            data_ws = data_template_wb[sheet_name]
            data_src_ws = data_source_wb[sheet_name] if data_source_wb and sheet_name in data_source_wb.sheetnames else None
            for mapping in [m for m in self.mappings if m.get('target_sheet') == sheet_name]:
                try:
                    self.execute_mapping(ws, mapping, data_ws, data_src_ws)
                except Exception as e:
                    failed_mappings.append(f"{sheet_name}/{mapping.get('type')}: {e}")

        data_template_wb.close()
        if data_source_wb:
            data_source_wb.close()

        save_path, _ = QFileDialog.getSaveFileName(self, "保存报告", "", "Excel文件 (*.xlsx)")
        if save_path:
            try:
                self.template_wb.save(save_path)
            except Exception as e:
                QMessageBox.critical(self, "保存失败", f"保存报告时出错：{e}")
                return

            self._image_streams.clear()

            if failed_mappings:
                QMessageBox.warning(self, "部分映射失败",
                    "以下映射未能成功执行：\n" + "\n".join(failed_mappings) +
                    "\n\n报告已保存，请手动检查这些区域。")
            else:
                QMessageBox.information(self, "完成", "报告已保存")

    def execute_mapping(self, ws, mapping, data_ws, data_src_ws=None):
        if mapping['type'] == 'data':
            self.apply_data_mapping(ws, mapping, data_src_ws)
        elif mapping['type'] == 'image':
            self.apply_image_mapping(ws, mapping)
        elif mapping['type'] == 'archive_shift_right':
            self.apply_archive_shift_right(ws, mapping, data_ws, data_src_ws)
        elif mapping['type'] == 'jmp':
            self.apply_jmp_mapping(ws, mapping)

    # ---------- 数据转换 ----------
    def _apply_transform(self, value, mapping):
        trans = mapping.get('transform', 'none')
        if trans == 'none' or value is None:
            return value
        try:
            if trans == 'div1000':
                return float(value) / 1000.0
            elif trans == 'strip_letters':
                if isinstance(value, str):
                    return re.sub(r'[a-zA-Z]+$', '', value)
                return value
            elif trans == 'custom':
                expr = mapping.get('transform_expr', '')
                if expr:
                    return eval(expr, {"x": value, "__builtins__": {}})
        except:
            pass
        return value

    def apply_data_mapping(self, ws, mapping, data_src_ws=None):
        t_min_row, t_min_col, t_max_row, t_max_col = mapping['target_range']
        src_ws = data_src_ws if data_src_ws else self.source_wb[mapping['source_sheet']]
        s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(mapping['source_range'])
        for i in range(t_max_row - t_min_row + 1):
            for j in range(t_max_col - t_min_col + 1):
                src_row = s_min_row + i
                src_col = s_min_col + j
                if src_row <= s_max_row and src_col <= s_max_col:
                    val = src_ws.cell(row=src_row, column=src_col).value
                    val = self._apply_transform(val, mapping)
                    ws.cell(row=t_min_row + i, column=t_min_col + j).value = val

    def remove_images_at_anchor(self, ws, anchor):
        to_remove = []
        for img in ws._images:
            if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                r = img.anchor._from.row + 1
                c = img.anchor._from.col + 1
                if f"{get_column_letter(c)}{r}" == anchor:
                    to_remove.append(img)
        for img in to_remove:
            ws._images.remove(img)

    def _process_image_data(self, img_bytes, rotation, target_w, target_h):
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        if rotation != 0:
            pil_img = pil_img.rotate(-rotation, expand=True)
        pil_img = pil_img.resize((target_w, target_h), PILImage.LANCZOS)
        out_stream = io.BytesIO()
        pil_img.save(out_stream, format='PNG')
        out_stream.seek(0)
        return out_stream

    def apply_image_mapping(self, ws, mapping):
        try:
            anchor = mapping['anchor_cell']
            col_width_chars = mapping.get('col_width_chars', 8.0)
            row_height_pts = mapping.get('row_height_pts', 15.0)
            rotation = mapping.get('rotation', 0.0)
            w_scale = mapping.get('width_scale', 1.0)
            h_scale = mapping.get('height_scale', 1.0)

            target_width = int(col_width_chars * COL_WIDTH_PX_PER_CHAR * w_scale)
            target_height = int(row_height_pts * ROW_HEIGHT_PX_PER_PT * h_scale)

            if 'image_path' in mapping:
                if not os.path.exists(mapping['image_path']):
                    raise FileNotFoundError(f"图片文件不存在: {mapping['image_path']}")
                with open(mapping['image_path'], 'rb') as f:
                    img_bytes = f.read()
                processed_stream = self._process_image_data(img_bytes, rotation, target_width, target_height)
                self._image_streams.append(processed_stream)

                self.remove_images_at_anchor(ws, anchor)
                new_img = OpenpyxlImage(processed_stream)
                new_img.anchor = anchor
                ws.add_image(new_img)
                return

            if 'image_bytes' not in mapping:
                raise RuntimeError("映射中缺少图片数据")
            img_bytes = mapping['image_bytes']
            processed_stream = self._process_image_data(img_bytes, rotation, target_width, target_height)
            self._image_streams.append(processed_stream)

            self.remove_images_at_anchor(ws, anchor)
            new_img = OpenpyxlImage(processed_stream)
            new_img.anchor = anchor
            ws.add_image(new_img)

        except Exception as e:
            QMessageBox.warning(self, "图片处理错误", f"图片写入失败：{e}")

    def apply_archive_shift_right(self, ws, mapping, data_ws, data_src_ws=None):
        min_row, min_col, max_row, max_col = mapping['block_range']
        header_rows = mapping.get('header_rows', 1)
        new_headers = mapping.get('new_headers', [])
        src_ws = data_src_ws if data_src_ws else self.source_wb[mapping['source_sheet']]
        s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(mapping['source_range'])

        merged_to_shift = []
        for mr in ws.merged_cells.ranges:
            m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(str(mr))
            if (m_min_row >= min_row and m_max_row <= max_row and
                m_min_col >= min_col and m_max_col <= max_col):
                merged_to_shift.append(str(mr))
        for mr in merged_to_shift:
            ws.unmerge_cells(mr)

        # 右移（纯数值）
        for col in range(max_col, min_col - 1, -1):
            for row in range(min_row, max_row + 1):
                src_cell = data_ws.cell(row=row, column=col)
                dst_cell = ws.cell(row=row, column=col + 1)
                dst_cell.value = src_cell.value
                self.copy_cell_style(ws.cell(row=row, column=col), dst_cell)

        for row in range(min_row, max_row + 1):
            ws.cell(row=row, column=min_col).value = None

        for mr in merged_to_shift:
            m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(mr)
            ws.merge_cells(f"{get_column_letter(m_min_col+1)}{m_min_row}:{get_column_letter(m_max_col+1)}{m_max_row}")

        # 填充新表头
        for i in range(header_rows):
            row_idx = min_row + i
            cell = ws.cell(row=row_idx, column=min_col)
            if i < len(new_headers):
                cell.value = new_headers[i]
            self.copy_cell_style(ws.cell(row=row_idx, column=min_col + 1), cell)

        # 填充新数据
        data_start_row = min_row + header_rows
        for i in range(max_row - data_start_row + 1):
            src_row = s_min_row + i
            if src_row > s_max_row:
                break
            target_row = data_start_row + i
            src_cell = src_ws.cell(row=src_row, column=s_min_col)
            ws.cell(row=target_row, column=min_col).value = src_cell.value
            self.copy_cell_style(ws.cell(row=target_row, column=min_col + 1),
                                 ws.cell(row=target_row, column=min_col))

    def apply_jmp_mapping(self, ws, mapping):
        anchor = mapping['anchor_cell']
        col_letter = ''.join(filter(str.isalpha, anchor))
        row_num = int(''.join(filter(str.isdigit, anchor)))
        start_row = row_num
        start_col = column_index_from_string(col_letter)

        headers = mapping['header_cols']  # 列表
        src_sheet_name = mapping['source_sheet']
        src_range = mapping['source_range']
        merge = mapping['merge_columns']

        src_ws = self.template_wb[src_sheet_name]
        s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(src_range)

        data_rows = s_max_row - s_min_row + 1
        data_cols = s_max_col - s_min_col + 1

        # 获取参考行样式（锚点上一行）
        ref_row = start_row - 1
        if ref_row < 1:
            ref_row = start_row  # 如果锚点在第1行，则参考自身

        if merge and data_cols > 1:
            values = []
            for r in range(s_min_row, s_max_row + 1):
                for c in range(s_min_col, s_max_col + 1):
                    val = src_ws.cell(row=r, column=c).value
                    values.append(val)
            row_count = len(values)
            for row_offset in range(row_count):
                current_row = start_row + row_offset
                # 写入表头
                for idx, header_text in enumerate(headers):
                    cell = ws.cell(row=current_row, column=start_col + idx)
                    cell.value = header_text
                # 写入数据
                val = values[row_offset]
                ws.cell(row=current_row, column=start_col + len(headers), value=val)
                # 复制格式（整行）
                for col_idx in range(start_col, start_col + len(headers) + 1):
                    ref_cell = ws.cell(row=ref_row, column=col_idx)
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    self.copy_cell_style(ref_cell, target_cell)
        else:
            row_count = data_rows
            target_data_cols = data_cols
            for row_offset in range(row_count):
                current_row = start_row + row_offset
                # 写入表头
                for idx, header_text in enumerate(headers):
                    cell = ws.cell(row=current_row, column=start_col + idx)
                    cell.value = header_text
                # 写入数据
                for c in range(target_data_cols):
                    src_val = src_ws.cell(row=s_min_row + row_offset, column=s_min_col + c).value
                    ws.cell(row=current_row, column=start_col + len(headers) + c, value=src_val)
                # 复制格式
                for col_idx in range(start_col, start_col + len(headers) + target_data_cols):
                    ref_cell = ws.cell(row=ref_row, column=col_idx)
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    self.copy_cell_style(ref_cell, target_cell)

    @staticmethod
    def copy_cell_style(src, dst):
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)

    # ==================== 配置保存 ====================
    def save_config(self):
        if not self.template_path:
            QMessageBox.warning(self, "提示", "请先打开模板文件")
            return
        clean_mappings = []
        for m in self.mappings:
            m_copy = m.copy()
            if 'image_bytes' in m_copy:
                m_copy['image_bytes'] = None
            clean_mappings.append(m_copy)

        config = {'template_file': self.template_path, 'mappings': clean_mappings}
        config_path = self.template_path.rsplit('.', 1)[0] + '_config.json'
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "完成", f"配置已保存至 {config_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"配置保存失败：{e}")



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path('app_icon.ico')))  # 图标也用 resource_path

    # 启动画面
    splash_pix = QPixmap(resource_path('splash.png'))   # ← 关键修改
    if not splash_pix.isNull():
        splash_pix = splash_pix.scaledToWidth(600, Qt.SmoothTransformation)
    else:
        # 备用：纯白背景 + 文字
        splash_pix = QPixmap(400, 200)
        splash_pix.fill(Qt.white)
        painter = QPainter(splash_pix)
        painter.setFont(QFont('Arial', 20))
        painter.drawText(splash_pix.rect(), Qt.AlignCenter, "自动报告工具")
        painter.end()

    splash = QSplashScreen(splash_pix)
    splash.show()
    splash.showMessage("正在启动...", Qt.AlignBottom | Qt.AlignCenter, Qt.black)
    app.processEvents()

    window = MainWindow()
    splash.finish(window)
    window.show()

    sys.exit(app.exec_())
