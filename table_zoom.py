"""表格缩放混入。

给 QTableWidget 增加 Ctrl+滚轮 / Mac 触控板双指捏合缩放：
同步缩放列宽、行高与单元格字体。事件 50ms 合并，避免快速操作卡顿。
"""
from PyQt5.QtCore import Qt, QEvent, QTimer


class TableZoomMixin:
    ZOOM_MIN = 0.4
    ZOOM_MAX = 3.0

    def enable_table_zoom(self, table):
        self._zoom_table = table
        self._zoom = 1.0
        self._zoom_base_w = {}
        self._zoom_base_h = {}
        self._zoom_items = []
        self._zoom_timer = QTimer(self)
        self._zoom_timer.setSingleShot(True)
        self._zoom_timer.timeout.connect(self._zoom_apply)
        viewport = table.viewport()
        viewport.installEventFilter(self)
        viewport.grabGesture(Qt.PinchGesture)

    def zoom_begin(self):
        """每次重新渲染表格前调用，清空已登记的字体 item"""
        self._zoom_items = []

    def zoom_size(self, col_idx, width, row_idx, height):
        """渲染尺寸阶段调用：记录基准尺寸，并按当前缩放倍数设置"""
        table = self._zoom_table
        if col_idx is not None and width:
            self._zoom_base_w[col_idx] = width
            table.setColumnWidth(col_idx, max(8, int(width * self._zoom)))
        if row_idx is not None and height:
            self._zoom_base_h[row_idx] = height
            table.setRowHeight(row_idx, max(8, int(height * self._zoom)))

    def zoom_track_item(self, row, col, base_font_size):
        """渲染单元格阶段调用：登记该 item 的基准字号，供缩放时重设"""
        self._zoom_items.append((row, col, base_font_size))

    def _zoom_request(self, factor):
        self._zoom = max(self.ZOOM_MIN, min(self.ZOOM_MAX, self._zoom * factor))
        self._zoom_timer.start(50)

    def _zoom_apply(self):
        table = self._zoom_table
        for c, w in self._zoom_base_w.items():
            table.setColumnWidth(c, max(8, int(w * self._zoom)))
        for r, h in self._zoom_base_h.items():
            table.setRowHeight(r, max(8, int(h * self._zoom)))
        for row, col, base in self._zoom_items:
            item = table.item(row, col)
            if item is None:
                continue
            font = item.font()
            font.setPointSize(max(5, int(base * self._zoom)))
            item.setFont(font)

    def eventFilter(self, obj, event):
        table = getattr(self, '_zoom_table', None)
        if table is not None and obj is table.viewport():
            if event.type() == QEvent.Wheel:
                if event.modifiers() & Qt.ControlModifier:
                    delta = event.angleDelta().y()
                    if delta:
                        self._zoom_request(1.15 if delta > 0 else 1 / 1.15)
                    return True
            elif event.type() == QEvent.Gesture:
                gesture = event.gesture(Qt.PinchGesture)
                if gesture is not None:
                    factor = gesture.scaleDelta()
                    if factor:
                        self._zoom_request(factor)
                    return True
        # 混入类在 MRO 中位于 Qt 类之后，super() 没有 eventFilter；
        # 事件过滤器默认放行（返回 False）即可
        return False
