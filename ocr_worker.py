"""OCR 后台工作线程：把批量 OCR 移到子线程，避免阻塞 UI。"""
from PyQt5.QtCore import QThread, pyqtSignal


class OCRWorker(QThread):
    """在子线程中运行 OCR 批量识别，通过信号汇报进度和结果。"""

    # 信号
    progress = pyqtSignal(int, int)       # current, total
    image_done = pyqtSignal(int, object)  # index, extracted_dict
    finished = pyqtSignal(list)           # all_results
    error = pyqtSignal(str)               # error_message

    def __init__(self, image_paths, roi=None, preprocess='none',
                 lang='ch', mode='first_number', labels=None, expr='',
                 roi_configs=None, parent=None):
        super().__init__(parent)
        self.image_paths = image_paths
        self.roi = roi
        self.preprocess = preprocess
        self.lang = lang
        self.mode = mode
        self.labels = labels
        self.expr = expr
        self.roi_configs = roi_configs

    def run(self):
        from ocr_engine import ocr_batch, ocr_batch_with_rois

        def on_progress(current, total):
            self.progress.emit(current, total)

        try:
            if self.roi_configs:
                results = ocr_batch_with_rois(
                    self.image_paths, roi_configs=self.roi_configs,
                    preprocess=self.preprocess, lang=self.lang,
                    mode=self.mode, labels=self.labels, expr=self.expr,
                    progress_callback=on_progress)
            else:
                results = ocr_batch(
                    self.image_paths, roi=self.roi,
                    preprocess=self.preprocess, lang=self.lang,
                    mode=self.mode, labels=self.labels, expr=self.expr,
                    progress_callback=on_progress)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))
