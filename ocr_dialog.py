"""OCR 切片量测配置与预览对话框。"""
import os
import re

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QPushButton,
    QFormLayout, QLineEdit, QComboBox, QDialogButtonBox, QListWidget,
    QListWidgetItem, QCheckBox, QGroupBox, QRadioButton, QButtonGroup,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QScrollArea,
    QProgressBar, QApplication, QFileDialog, QSplitter, QMessageBox,
    QSpinBox,
)
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QFont
from PyQt5.QtCore import Qt, QRect, QPoint, QSize, pyqtSignal

from constants import (
    OCR_DEFAULT_LANG, OCR_SUPPORTED_LANGS,
    OCR_EXTRACTION_MODES, OCR_PREPROCESS_OPTIONS,
    OCR_MAX_IMAGE_DIMENSION_FOR_PREVIEW, OCR_PREVIEW_COUNT,
)
from ocr_engine import ocr_available


# ---------- ROI 选择器控件 ----------
class ROISelectorWidget(QWidget):
    """图像预览 + 鼠标拖拽框选 ROI 区域。"""
    roi_changed = pyqtSignal(object)  # (x, y, w, h) or None

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pixmap = None
        self._display_pixmap = None
        self._offset_x = 0
        self._offset_y = 0
        self._scale = 1.0
        self._start = None
        self._end = None
        self._roi = None  # (x, y, w, h) in original image pixel space
        self.setMinimumSize(400, 300)
        self.setMouseTracking(True)
        self.setCursor(Qt.CrossCursor)
        self.setStyleSheet("background: #f0f0f0; border: 1px solid #ccc;")

    def set_image(self, path):
        """加载图像并缩放到适合控件大小。"""
        pix = QPixmap(path)
        if pix.isNull():
            self._pixmap = None
            self._display_pixmap = None
            self.update()
            return
        self._pixmap = pix
        self._fit_to_widget()
        self._roi = None
        self._start = None
        self._end = None
        self.update()

    def _fit_to_widget(self):
        if self._pixmap is None:
            return
        w = self.width() - 20
        h = self.height() - 20
        if w <= 0 or h <= 0:
            return
        scaled = self._pixmap.scaled(
            QSize(w, h), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self._display_pixmap = scaled
        self._scale = self._pixmap.width() / scaled.width()
        self._offset_x = (self.width() - scaled.width()) // 2
        self._offset_y = (self.height() - scaled.height()) // 2

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_to_widget()
        self.update()

    def _to_original(self, px, py):
        """显示坐标 → 原始图像坐标"""
        ox = (px - self._offset_x) * self._scale
        oy = (py - self._offset_y) * self._scale
        return int(ox), int(oy)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#f0f0f0"))

        if self._display_pixmap is None:
            painter.setPen(QColor("#999"))
            painter.drawText(self.rect(), Qt.AlignCenter, "（无图像）")
            return

        painter.drawPixmap(self._offset_x, self._offset_y,
                           self._display_pixmap)

        if self._start and self._end:
            x1, y1 = self._start.x(), self._start.y()
            x2, y2 = self._end.x(), self._end.y()
            r = QRect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
            pen = QPen(QColor("#00ff00"), 2, Qt.DashLine)
            painter.setPen(pen)
            painter.setBrush(QColor(0, 255, 0, 30))
            painter.drawRect(r)

        # 绘制已有 ROI
        if self._roi and not self._start:
            x, y, w, h = self._roi
            sx = int(x / self._scale) + self._offset_x
            sy = int(y / self._scale) + self._offset_y
            sw = int(w / self._scale)
            sh = int(h / self._scale)
            pen = QPen(QColor("#ff6600"), 2, Qt.DashLine)
            painter.setPen(pen)
            painter.setBrush(QColor(255, 102, 0, 30))
            painter.drawRect(QRect(sx, sy, sw, sh))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._display_pixmap:
            self._start = event.pos()
            self._end = event.pos()
            self.update()

    def mouseMoveEvent(self, event):
        if self._start and self._display_pixmap:
            self._end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._start:
            self._end = event.pos()
            x1, y1 = self._start.x(), self._start.y()
            x2, y2 = self._end.x(), self._end.y()
            r = QRect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
            if r.width() > 5 and r.height() > 5:
                ox, oy = self._to_original(r.x(), r.y())
                ow, oh = int(r.width() * self._scale), int(r.height() * self._scale)
                self._roi = (ox, oy, ow, oh)
                self.roi_changed.emit(self._roi)
            else:
                self._roi = None
                self.roi_changed.emit(None)
            self._start = None
            self._end = None
            self.update()

    def get_roi(self):
        return self._roi

    def set_roi(self, roi):
        self._roi = roi
        self.update()

    def clear_roi(self):
        self._roi = None
        self._start = None
        self._end = None
        self.roi_changed.emit(None)
        self.update()


# ---------- OCR 配置与预览对话框 ----------
class OCRSetupDialog(QDialog):
    """OCR 切片量测识别配置对话框。

    使用方式:
        dlg = OCRSetupDialog(parent=parent)
        if dlg.exec_() == QDialog.Accepted:
            mapping = dlg.get_ocr_mapping()
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("切片量测 OCR 识别")
        self.setMinimumSize(1100, 750)

        self._image_folder = ""
        self._all_images = []       # 所有匹配的图片绝对路径
        self._selected_images = []  # 用户勾选的图片路径
        self._preview_results = []  # 预览 OCR 结果
        self._labels = []           # 多值模式的标签

        self._init_ui()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)

        # ── 模式选择 ──
        mode_group = QGroupBox("识别模式")
        mode_layout = QHBoxLayout(mode_group)
        self.mode_group = QButtonGroup(self)
        self.radio_single = QRadioButton("单值模式（每图提取一个数值）")
        self.radio_labeled = QRadioButton("多值标注模式（每图提取多个标注值）")
        self.mode_group.addButton(self.radio_single, 0)
        self.mode_group.addButton(self.radio_labeled, 1)
        self.radio_single.setChecked(True)
        mode_layout.addWidget(self.radio_single)
        mode_layout.addWidget(self.radio_labeled)
        mode_layout.addStretch()
        self.mode_group.buttonClicked.connect(self._on_mode_changed)
        main_layout.addWidget(mode_group)

        # ── 上半部分：图像源 + ROI ──
        top_split = QSplitter(Qt.Horizontal)

        # 左：图像源
        src_widget = QWidget()
        src_layout = QVBoxLayout(src_widget)
        src_layout.setContentsMargins(0, 0, 0, 0)

        # 文件夹选择
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("切片文件夹:"))
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("选择 X-section 文件夹…")
        self.folder_edit.setReadOnly(True)
        folder_layout.addWidget(self.folder_edit, 1)
        btn_browse = QPushButton("浏览…")
        btn_browse.clicked.connect(self._browse_folder)
        folder_layout.addWidget(btn_browse)
        src_layout.addLayout(folder_layout)

        # 文件名过滤
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("文件名过滤:"))
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("如 _副本 或 -1（留空=全部）")
        self.filter_edit.textChanged.connect(self._refresh_image_list)
        filter_layout.addWidget(self.filter_edit, 1)
        src_layout.addLayout(filter_layout)

        # 子文件夹选择
        subf_layout = QHBoxLayout()
        subf_layout.addWidget(QLabel("子文件夹:"))
        self.subfolder_combo = QComboBox()
        self.subfolder_combo.addItem("（请先选择文件夹）")
        self.subfolder_combo.currentTextChanged.connect(self._refresh_image_list)
        subf_layout.addWidget(self.subfolder_combo, 1)
        src_layout.addLayout(subf_layout)

        # 图像列表
        self.image_list = QListWidget()
        self.image_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.image_list.itemSelectionChanged.connect(self._on_image_selection_changed)
        src_layout.addWidget(self.image_list, 1)

        # 选择操作按钮
        sel_btn_layout = QHBoxLayout()
        btn_all = QPushButton("全选")
        btn_all.clicked.connect(lambda: self._select_all_images(True))
        btn_none = QPushButton("取消全选")
        btn_none.clicked.connect(lambda: self._select_all_images(False))
        sel_btn_layout.addWidget(btn_all)
        sel_btn_layout.addWidget(btn_none)
        sel_btn_layout.addStretch()
        self.count_label = QLabel("共 0 张")
        sel_btn_layout.addWidget(self.count_label)
        src_layout.addLayout(sel_btn_layout)

        top_split.addWidget(src_widget)

        # 右：ROI + 预览图像
        roi_widget = QWidget()
        roi_layout = QVBoxLayout(roi_widget)
        roi_layout.setContentsMargins(0, 0, 0, 0)

        roi_group = QGroupBox("ROI 区域选择")
        roi_group_layout = QVBoxLayout(roi_group)
        self.roi_viewer = ROISelectorWidget()
        self.roi_viewer.roi_changed.connect(self._on_roi_changed)
        roi_group_layout.addWidget(self.roi_viewer, 1)

        roi_info_layout = QHBoxLayout()
        self.roi_label = QLabel("ROI: 未选择")
        roi_info_layout.addWidget(self.roi_label)
        btn_clear_roi = QPushButton("清除 ROI")
        btn_clear_roi.clicked.connect(self.roi_viewer.clear_roi)
        roi_info_layout.addWidget(btn_clear_roi)
        roi_info_layout.addStretch()
        roi_group_layout.addLayout(roi_info_layout)

        roi_layout.addWidget(roi_group, 1)

        top_split.addWidget(roi_widget)
        top_split.setSizes([400, 600])
        main_layout.addWidget(top_split, 1)

        # ── 标签配置（仅多值模式显示）──
        self.label_group = QGroupBox("量测标签配置（多值模式）")
        label_layout = QHBoxLayout(self.label_group)
        label_layout.addWidget(QLabel("标签列表（逗号分隔）:"))
        self.labels_edit = QLineEdit()
        self.labels_edit.setPlaceholderText("No.2, No.3, No.4, No.5")
        self.labels_edit.textChanged.connect(self._on_labels_changed)
        label_layout.addWidget(self.labels_edit, 1)
        self.label_group.setVisible(False)
        main_layout.addWidget(self.label_group)

        # ── OCR 设置 ──
        ocr_settings = QGroupBox("OCR 设置")
        settings_layout = QHBoxLayout(ocr_settings)
        settings_layout.addWidget(QLabel("识别语言:"))
        self.lang_combo = QComboBox()
        for lang in OCR_SUPPORTED_LANGS:
            self.lang_combo.addItem({'ch': '中文+数字', 'en': '英文+数字'}[lang], lang)
        settings_layout.addWidget(self.lang_combo)

        settings_layout.addWidget(QLabel("预处理:"))
        self.preprocess_combo = QComboBox()
        pre_names = {'none': '无', 'grayscale': '灰度化', 'otsu': 'OTSU二值化',
                     'stretch': '直方图拉伸', 'stretch_invert': '拉伸+反转'}
        for opt in OCR_PREPROCESS_OPTIONS:
            self.preprocess_combo.addItem(pre_names[opt], opt)
        settings_layout.addWidget(self.preprocess_combo)

        settings_layout.addWidget(QLabel("提取方式:"))
        self.extraction_combo = QComboBox()
        # 初始为单值模式的提取选项（多值选项在切换模式时重建）
        self.extraction_combo.addItem("首个数值", "first_number")
        self.extraction_combo.addItem("全部数值", "all_numbers")
        self.extraction_combo.addItem("自定义", "custom")
        settings_layout.addWidget(self.extraction_combo)

        settings_layout.addWidget(QLabel("自定义表达式:"))
        self.expr_edit = QLineEdit()
        self.expr_edit.setPlaceholderText("如 round(x,2)（x=整图识别文本）")
        self.expr_edit.setEnabled(False)
        settings_layout.addWidget(self.expr_edit)

        settings_layout.addStretch()
        self.extraction_combo.currentTextChanged.connect(self._on_extraction_changed)
        main_layout.addWidget(ocr_settings)

        # ── 预览结果 ──
        preview_group = QGroupBox("预览识别结果")
        preview_layout = QVBoxLayout(preview_group)

        self.preview_table = QTableWidget()
        self.preview_table.setMinimumHeight(120)
        preview_layout.addWidget(self.preview_table)

        preview_btn_layout = QHBoxLayout()
        self.btn_run_preview = QPushButton("运行预览 OCR")
        self.btn_run_preview.clicked.connect(self._run_preview_ocr)
        self.btn_run_preview.setEnabled(False)
        preview_btn_layout.addWidget(self.btn_run_preview)

        self.preview_progress = QProgressBar()
        self.preview_progress.setVisible(False)
        preview_btn_layout.addWidget(self.preview_progress, 1)
        preview_btn_layout.addStretch()
        preview_layout.addLayout(preview_btn_layout)

        main_layout.addWidget(preview_group)

        # ── 底部 ──
        # OCR 可用性提示
        if not ocr_available():
            warn = QLabel("⚠ PaddleOCR 未安装，请运行: pip install paddlepaddle paddleocr")
            warn.setStyleSheet("color: #b00020; font-weight: bold;")
            main_layout.addWidget(warn)
            self.btn_run_preview.setEnabled(False)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)

    # ── 模式切换 ──
    def _on_mode_changed(self, btn):
        is_labeled = (self.mode_group.id(btn) == 1)
        self.label_group.setVisible(is_labeled)
        # 更新提取方式选项
        self.extraction_combo.clear()
        if is_labeled:
            self.extraction_combo.addItem("按标签", "labeled")
            self.extraction_combo.addItem("全部数值", "all_numbers")
            self.extraction_combo.addItem("自定义", "custom")
        else:
            self.extraction_combo.addItem("首个数值", "first_number")
            self.extraction_combo.addItem("全部数值", "all_numbers")
            self.extraction_combo.addItem("自定义", "custom")

    def _on_extraction_changed(self):
        """自定义提取模式时启用表达式输入框"""
        if hasattr(self, 'expr_edit'):
            self.expr_edit.setEnabled(
                self.extraction_combo.currentData() == 'custom')

    # ── 文件夹浏览 ──
    def _browse_folder(self):
        start = self.folder_edit.text() or ""
        folder = QFileDialog.getExistingDirectory(self, "选择切片文件夹", start)
        if not folder:
            return
        self._image_folder = folder
        self.folder_edit.setText(folder)
        # 扫描子文件夹
        self.subfolder_combo.clear()
        self.subfolder_combo.addItem("（全部子文件夹）", "")
        try:
            for entry in sorted(os.listdir(folder)):
                full = os.path.join(folder, entry)
                if os.path.isdir(full) and not entry.startswith('.'):
                    self.subfolder_combo.addItem(entry, full)
        except OSError:
            pass
        self._refresh_image_list()

    # ── 图像列表刷新 ──
    def _refresh_image_list(self):
        """根据文件夹、子文件夹过滤、文件名过滤重建图像列表"""
        folder = self._image_folder
        if not folder or not os.path.isdir(folder):
            self._all_images = []
            self.image_list.clear()
            self.count_label.setText("共 0 张")
            self.btn_run_preview.setEnabled(False)
            return

        subfolder = self.subfolder_combo.currentData()
        search_dirs = [subfolder] if subfolder else [
            os.path.join(folder, d) for d in os.listdir(folder)
            if os.path.isdir(os.path.join(folder, d)) and not d.startswith('.')
        ]
        if not search_dirs:
            search_dirs = [folder]

        filt = self.filter_edit.text().strip()
        images = []
        for sd in search_dirs:
            try:
                for f in sorted(os.listdir(sd)):
                    if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                        if filt and filt not in f:
                            continue
                        images.append(os.path.join(sd, f))
            except OSError:
                continue

        self._all_images = images
        self.image_list.clear()
        for img in images:
            item = QListWidgetItem(os.path.relpath(img, folder))
            item.setData(Qt.UserRole, img)
            item.setSelected(True)  # 默认全选
            self.image_list.addItem(item)

        self.count_label.setText(f"共 {len(images)} 张")
        self.btn_run_preview.setEnabled(len(images) > 0 and ocr_available())

        # 加载第一张图到 ROI 预览
        if images:
            self.roi_viewer.set_image(images[0])

    # ── 图像选择 ──
    def _select_all_images(self, select):
        for i in range(self.image_list.count()):
            self.image_list.item(i).setSelected(select)

    def _on_image_selection_changed(self):
        items = self.image_list.selectedItems()
        self._selected_images = [it.data(Qt.UserRole) for it in items]
        self.count_label.setText(
            f"已选 {len(self._selected_images)}/{len(self._all_images)} 张")

        # 切换 ROI 预览为第一张选中的图
        if self._selected_images:
            self.roi_viewer.set_image(self._selected_images[0])

    # ── ROI ──
    def _on_roi_changed(self, roi):
        if roi is None:
            self.roi_label.setText("ROI: 未选择（将对整张图做 OCR）")
        else:
            self.roi_label.setText(
                f"ROI: x={roi[0]}, y={roi[1]}, w={roi[2]}, h={roi[3]}")

    # ── 标签 ──
    def _on_labels_changed(self, text):
        parts = re.split(r'[,，、\s]+', text.strip())
        self._labels = [p for p in parts if p]

    # ── 预览 OCR ──
    def _run_preview_ocr(self):
        if not self._selected_images:
            return

        # 取前 N 张做预览
        preview_images = self._selected_images[:OCR_PREVIEW_COUNT]
        roi = self.roi_viewer.get_roi()
        preprocess = self.preprocess_combo.currentData()
        lang = self.lang_combo.currentData()
        mode = self.extraction_combo.currentData()
        labels = self._labels if mode == 'labeled' else None
        expr = self.expr_edit.text() if mode == 'custom' else ''

        self.btn_run_preview.setEnabled(False)
        self.preview_progress.setVisible(True)
        self.preview_progress.setRange(0, len(preview_images))

        from ocr_worker import OCRWorker
        self._preview_worker = OCRWorker(
            preview_images, roi=roi, preprocess=preprocess, lang=lang,
            mode=mode, labels=labels, expr=expr,
        )
        self._preview_worker.progress.connect(self._on_preview_progress)
        self._preview_worker.finished.connect(self._on_preview_finished)
        self._preview_worker.error.connect(self._on_preview_error)
        self._preview_worker.start()

    def _on_preview_progress(self, current, total):
        self.preview_progress.setValue(current)
        QApplication.processEvents()

    def _on_preview_finished(self, results):
        self._preview_results = results
        self._display_preview_results(results)
        self.preview_progress.setVisible(False)
        self.btn_run_preview.setEnabled(True)

    def _on_preview_error(self, msg):
        QMessageBox.critical(self, "OCR 错误", f"预览 OCR 失败：{msg}")
        self.preview_progress.setVisible(False)
        self.btn_run_preview.setEnabled(True)

    def _display_preview_results(self, results):
        """在预览表中显示 OCR 结果"""
        mode = self.extraction_combo.currentData()
        table = self.preview_table

        if mode == 'labeled' and self._labels:
            # 表头: 图像 | 标签1 | 标签2 | ...
            table.setColumnCount(1 + len(self._labels))
            headers = ["图像"] + self._labels
            table.setHorizontalHeaderLabels(headers)
        elif mode == 'all_numbers':
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(["图像", "识别到的数值"])
        else:
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(["图像", "识别值"])

        table.setRowCount(len(results))
        for i, r in enumerate(results):
            img_name = r.get('_image', '?')
            item = QTableWidgetItem(img_name)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 0, item)

            if '_error' in r:
                err_item = QTableWidgetItem(f"[错误] {r['_error']}")
                err_item.setForeground(QColor("#b00020"))
                err_item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(i, 1, err_item)
                continue

            if mode == 'labeled' and self._labels:
                for j, label in enumerate(self._labels):
                    v = r.get(label)
                    text = str(v) if v is not None else "[N/A]"
                    cell = QTableWidgetItem(text)
                    cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                    if v is None:
                        cell.setForeground(QColor("#999"))
                    table.setItem(i, 1 + j, cell)
            elif mode == 'all_numbers':
                vals = r.get('values', [])
                cell = QTableWidgetItem(str(vals) if vals else "[无]")
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                table.setItem(i, 1, cell)
            else:
                v = r.get('value')
                text = str(v) if v is not None else "[未识别]"
                cell = QTableWidgetItem(text)
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                if v is None:
                    cell.setForeground(QColor("#b00020"))
                table.setItem(i, 1, cell)

        table.resizeColumnsToContents()

    # ── 确认 ──
    def _on_accept(self):
        if not self._selected_images:
            QMessageBox.warning(self, "提示", "请至少选择一张切片图像。")
            return
        self.accept()

    # ── 公共接口 ──
    def get_ocr_mapping(self):
        """返回 OCR 映射 dict，供 main.py 添加到 self.mappings。"""
        mode = self.extraction_combo.currentData()
        roi = self.roi_viewer.get_roi()
        preprocess = self.preprocess_combo.currentData()
        lang = self.lang_combo.currentData()

        mapping = {
            'type': 'ocr',
            'mode': mode,
            'target_sheet': '',  # 由 main.py 填充
            'image_folder': self._image_folder,
            'image_list': [os.path.relpath(p, self._image_folder)
                           for p in self._selected_images],
            'roi': roi,
            'preprocess': preprocess,
            'lang': lang,
        }

        if mode == 'labeled':
            mapping['labels'] = self._labels
        elif mode == 'custom':
            mapping['expr'] = self.expr_edit.text()

        return mapping

    def get_selected_count(self):
        return len(self._selected_images)

    def get_labels(self):
        return self._labels
