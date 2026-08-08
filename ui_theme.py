"""Apple 风格 UI 设计体系（移植自 MC_LogAnalysis 的界面设计经验）。

使用方式:
    app.setStyleSheet(APPLE_QSS)      # 全局生效（主窗口 + 所有对话框）
    w, h = window_target_size()       # 主窗口按屏幕可用区域 80% 动态计算
    widget.setProperty("card", True)  # 白色圆角卡片容器

要点（来自项目记忆）:
    - 背景 #F5F5F7、卡片白、主色 #007AFF
    - 窗口尺寸 = 屏幕可用区域 80%（最小 900×620）
    - 启动画面与主窗口同尺寸，淡出 300ms 衔接（见 main.py）
    - AA_EnableHighDpiScaling 必须在 QApplication 创建前设置（main.py 已做）
"""
from PyQt5.QtWidgets import QApplication

# ── 色彩（Apple 设计体系）──────────────────────────────────────
C_BG       = "#F5F5F7"   # 窗口背景
C_CARD     = "#FFFFFF"   # 卡片底色
C_PRIME    = "#007AFF"   # 主色调（Apple Blue）
C_PRIME_H  = "#0062CC"   # 主色调 hover
C_TEXT     = "#1D1D1F"   # 主文字
C_SUB      = "#86868B"   # 辅助文字
C_BORDER   = "#E5E5EA"   # 边框/分割线
C_INPUT_BG = "#F9F9F9"   # 输入框底板
C_GREEN    = "#34C759"
C_RED      = "#FF3B30"

# 字体（macOS → SF Pro / Windows → Segoe UI）
FONT_FAMILY = (
    '"Helvetica Neue", "PingFang SC", "Segoe UI", "Microsoft YaHei", sans-serif'
)
FONT_MONO = '"Menlo", "Consolas", "Cascadia Code", "SF Mono", monospace'

# 间距
GAP_SECTION = 10        # 卡片间距
GAP_ROW     = 6         # 行内控件间距
GAP_INNER   = 6         # 标签-控件间距
CARD_PAD    = 12        # 卡片内边距
RADIUS      = 8         # 圆角

# 窗口尺寸
WINDOW_SCREEN_RATIO = 0.8   # 主窗口占屏幕可用区域百分比
WINDOW_MIN_SIZE = (900, 620)
WINDOW_DEFAULT_SIZE = (1400, 850)


def window_target_size(ratio=WINDOW_SCREEN_RATIO):
    """按当前屏幕可用区域计算主窗口目标尺寸（默认 80%，最小 900×620）。"""
    screen = QApplication.primaryScreen()
    if screen is None:
        return WINDOW_DEFAULT_SIZE
    geo = screen.availableGeometry()
    w = max(WINDOW_MIN_SIZE[0], int(geo.width() * ratio))
    h = max(WINDOW_MIN_SIZE[1], int(geo.height() * ratio))
    return w, h


# ── 全局样式表 ──────────────────────────────────────────────────
APPLE_QSS = f"""
/* ─── 全局 ─── */
QMainWindow, QDialog {{
    background-color: {C_BG};
}}
QWidget {{
    font-family: {FONT_FAMILY};
    font-size: 13px;
    color: {C_TEXT};
}}

/* ─── 卡片容器 ─── */
QWidget[card="true"] {{
    background-color: {C_CARD};
    border: 1px solid {C_BORDER};
    border-radius: {RADIUS}px;
}}

/* ─── 标签 ─── */
QLabel[heading="true"] {{
    font-size: 15px;
    font-weight: bold;
    color: {C_TEXT};
    padding: 0;
}}
QLabel[subtitle="true"] {{
    font-size: 12px;
    color: {C_SUB};
}}

/* ─── 主按钮（实心蓝） ─── */
QPushButton[primary="true"] {{
    background-color: {C_PRIME};
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 8px 24px;
    font-weight: bold;
    font-size: 13px;
}}
QPushButton[primary="true"]:hover {{
    background-color: {C_PRIME_H};
}}
QPushButton[primary="true"]:pressed {{
    background-color: #0055AA;
}}

/* ─── 次按钮（浅灰底） ─── */
QPushButton[secondary="true"] {{
    background-color: #F0F0F2;
    color: {C_TEXT};
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-size: 13px;
}}
QPushButton[secondary="true"]:hover {{
    background-color: #E4E4E8;
}}
QPushButton[secondary="true"]:pressed {{
    background-color: #D8D8DC;
}}

/* ─── 文字按钮（蓝色链接风） ─── */
QPushButton[link="true"] {{
    background: transparent;
    color: {C_PRIME};
    border: none;
    padding: 6px 12px;
    font-size: 13px;
}}
QPushButton[link="true"]:hover {{
    color: {C_PRIME_H};
    text-decoration: underline;
}}

/* ─── 输入框 ─── */
QLineEdit, QDoubleSpinBox, QSpinBox {{
    background-color: {C_INPUT_BG};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    font-size: 13px;
    color: {C_TEXT};
}}
QLineEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus {{
    border: 1.5px solid {C_PRIME};
    background-color: #FFFFFF;
}}
QLineEdit[readOnly="true"] {{
    background-color: {C_BG};
    color: {C_SUB};
}}

/* ─── 下拉框 ─── */
QComboBox {{
    background-color: {C_INPUT_BG};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    font-size: 13px;
    min-width: 160px;
}}
QComboBox:focus {{
    border: 1.5px solid {C_PRIME};
}}
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    border: none;
    border-left: 1px solid {C_BORDER};
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
}}
QComboBox QAbstractItemView {{
    background-color: #FFFFFF;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    selection-background-color: {C_PRIME};
    selection-color: #FFFFFF;
    padding: 4px;
    outline: none;
}}

/* ─── 复选/单选 ─── */
QCheckBox, QRadioButton {{
    font-size: 13px;
    color: {C_TEXT};
    spacing: 6px;
}}

/* ─── 进度条 ─── */
QProgressBar {{
    background-color: #E8E8ED;
    border: none;
    border-radius: 4px;
    height: 6px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    background-color: {C_PRIME};
    border-radius: 4px;
}}

/* ─── 文本区 ─── */
QTextEdit, QPlainTextEdit {{
    background-color: #FAFAFA;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 10px;
    font-family: {FONT_MONO};
    font-size: 12px;
    color: {C_TEXT};
}}

/* ─── 列表 / 表格 ─── */
QListWidget {{
    background-color: #FFFFFF;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    outline: none;
}}
QListWidget::item {{
    padding: 5px 8px;
    border-radius: 4px;
}}
QListWidget::item:selected {{
    background-color: {C_PRIME};
    color: #FFFFFF;
}}
QTableWidget {{
    background-color: #FFFFFF;
    alternate-background-color: #F7F7F7;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    gridline-color: #E5E5EA;
    outline: none;
}}
QTableWidget::item:selected {{
    background-color: {C_PRIME};
    color: #FFFFFF;
}}

/* ─── 滚动条 ─── */
QScrollBar:vertical {{
    background: transparent;
    width: 8px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: #D0D0D6;
    border-radius: 4px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background: #B0B0B6;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 8px;
    margin: 0;
}}
QScrollBar::handle:horizontal {{
    background: #D0D0D6;
    border-radius: 4px;
    min-width: 30px;
}}
QScrollBar::handle:horizontal:hover {{
    background: #B0B0B6;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* ─── 菜单栏 ─── */
QMenuBar {{
    background-color: #FFFFFF;
    border-bottom: 1px solid {C_BORDER};
    padding: 2px 0;
}}
QMenuBar::item {{
    padding: 6px 14px;
    border-radius: 5px;
}}
QMenuBar::item:selected {{
    background-color: #F0F0F2;
}}
QMenu {{
    background-color: #FFFFFF;
    border: 1px solid {C_BORDER};
    border-radius: 8px;
    padding: 6px 0;
}}
QMenu::item {{
    padding: 7px 32px 7px 20px;
}}
QMenu::item:selected {{
    background-color: {C_PRIME};
    color: #FFFFFF;
    border-radius: 4px;
}}

/* ─── 状态栏 ─── */
QStatusBar {{
    background-color: #FFFFFF;
    border-top: 1px solid {C_BORDER};
    font-size: 12px;
    color: {C_SUB};
}}

/* ─── 提示框 ─── */
QToolTip {{
    background-color: #FFFFFF;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
    color: {C_TEXT};
}}
"""
