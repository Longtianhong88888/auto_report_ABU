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
OCR_PREVIEW_COUNT = 3  # 预览时默认跑前 N 张图
OCR_MAX_IMAGE_SIDE = 2048  # OCR 识别前原图最长边上限（像素），超大图先等比压缩再识别

# 图片插入输出统一用 JPEG（比 PNG 体积小、编码快，降低输出文件与内存占用）
IMAGE_JPEG_QUALITY = 90          # 默认 JPEG 质量（PNG→JPEG、小 JPEG）
IMAGE_JPEG_QUALITY_LARGE = 85    # 大图 JPEG 重编码质量（>200KB 的原 JPEG，测试 q=85 肉眼无差异）
IMAGE_JPEG_LARGE_THRESHOLD_KB = 200  # 重编码触发阈值
IMAGE_DOWNSAMPLE_MAX_SIDE = 2048  # 图片处理降采样最长边（像素），超大图先缩至此再 resize 到目标尺寸，大幅降低内存峰值
IMAGE_THUMB_MAX_SIDE = 512  # 图片预览缩略图最长边（像素），JPEG 缩略图避免全图解码卡顿
