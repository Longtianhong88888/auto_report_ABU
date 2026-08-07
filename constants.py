# ================== 统一尺寸常量 ==================
COL_WIDTH_PX_PER_CHAR = 7.5
ROW_HEIGHT_PX_PER_PT = 1.333
DEFAULT_COL_WIDTH_CHARS = 8.0
DEFAULT_ROW_HEIGHT_PTS = 15.0
DEFAULT_COL_WIDTH_PX = int(DEFAULT_COL_WIDTH_CHARS * COL_WIDTH_PX_PER_CHAR)
DEFAULT_ROW_HEIGHT_PX = int(DEFAULT_ROW_HEIGHT_PTS * ROW_HEIGHT_PX_PER_PT)

# 预览渲染上限（行数按内容实际范围计算，此仅为防止病态超大表）
MAX_PREVIEW_ROWS = 10000
MAX_PREVIEW_COLS = 100

# ================== 用户ID授权 ==================
DEFAULT_AUTHORIZED_USER_IDS = [
    'G1655895',
]

# 管理员工号与授权密码
ADMIN_USER_ID = 'G1659304'
ADMIN_AUTH_PASSWORD = 'Zy1659304'

# ================== OCR 常量 ==================
OCR_DEFAULT_LANG = 'ch'
OCR_DEFAULT_CONFIDENCE_THRESHOLD = 0.5
OCR_SUPPORTED_LANGS = ['ch', 'en']
OCR_EXTRACTION_MODES = ['first_number', 'labeled', 'all_numbers', 'custom']
OCR_PREPROCESS_OPTIONS = ['none', 'grayscale', 'otsu', 'stretch', 'stretch_invert']
OCR_MAX_IMAGE_DIMENSION_FOR_PREVIEW = 1200
OCR_PREVIEW_COUNT = 3  # 预览时默认跑前 N 张图
