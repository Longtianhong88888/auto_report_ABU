"""OCR 引擎：PaddleOCR 封装、图像预处理、量测数值提取。

支持两种识别模式：
- single:     每张图提取最早出现的数值 → 单值
- labeled:    按标签（如 No.2, No.3）搜索标注值 → 多值字典
- all_numbers: 提取图中所有数值
- custom:     用 safe_eval 自定义表达式处理 OCR 原始文本
"""
import os
import io
import re
import logging
from functools import lru_cache

logger = logging.getLogger(__name__)

# ---------- 懒加载 PaddleOCR 单例 ----------
_ocr_instance = None
_ocr_init_error = None


def _get_ocr(lang='ch'):
    """懒加载 PaddleOCR 单例。

    模型查找顺序：
    1. 内置 paddleocr_models/official_models/{det,rec}/（离线打包）
    2. 用户目录 ~/.paddlex/（开发环境自动下载）
    """
    global _ocr_instance, _ocr_init_error
    if _ocr_instance is not None:
        return _ocr_instance
    if _ocr_init_error is not None:
        raise _ocr_init_error
    try:
        import sys
        # 确定内置模型路径
        bundled = None
        if hasattr(sys, '_MEIPASS'):
            bundled = os.path.join(sys._MEIPASS, 'paddleocr_models')
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            bundled = os.path.join(script_dir, 'paddleocr_models')

        # 检查内置模型
        det_dir = None
        rec_dir = None
        if bundled:
            official = os.path.join(bundled, 'official_models')
            if os.path.isdir(official):
                # 找到检测模型目录（第一个匹配的）
                for d in os.listdir(official):
                    dpath = os.path.join(official, d)
                    if os.path.isdir(dpath):
                        if 'det' in d.lower() and det_dir is None:
                            det_dir = dpath
                        elif 'rec' in d.lower() and rec_dir is None:
                            rec_dir = dpath

        from paddleocr import PaddleOCR
        kwargs = dict(
            lang=lang,
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )
        if det_dir:
            kwargs['text_detection_model_dir'] = det_dir
        if rec_dir:
            kwargs['text_recognition_model_dir'] = rec_dir

        _ocr_instance = PaddleOCR(**kwargs)
        return _ocr_instance
    except ImportError as e:
        _ocr_init_error = ImportError(
            "PaddleOCR 未安装。请运行: pip install paddlepaddle paddleocr\n"
            f"原始错误: {e}")
        raise _ocr_init_error
    except Exception as e:
        _ocr_init_error = RuntimeError(f"PaddleOCR 初始化失败: {e}")
        raise _ocr_init_error


def ocr_available():
    """检查 PaddleOCR 是否可用（不触发模型下载）"""
    try:
        import paddleocr  # noqa: F401
        return True
    except ImportError:
        return False


def _run_ocr(ocr, arr):
    """兼容 paddleocr 2.x 的 ocr() 与 3.x 的 predict() 调用"""
    if hasattr(ocr, 'ocr'):
        try:
            return ocr.ocr(arr)
        except (AttributeError, TypeError):
            pass
    if hasattr(ocr, 'predict'):
        return ocr.predict(arr)
    raise RuntimeError("PaddleOCR 实例不支持 ocr()/predict() 调用")


def _normalize_ocr_result(result):
    """把不同版本返回的单条结果归一化为 {text, confidence, bbox}"""
    out = []
    # 3.x OCRResult 对象
    if hasattr(result, 'rec_texts'):
        texts = result.rec_texts or []
        scores = result.rec_scores or []
        polys = result.rec_polys if hasattr(result, 'rec_polys') else []
        for i, (text, score) in enumerate(zip(texts, scores)):
            bbox = polys[i] if i < len(polys) else []
            out.append({'text': text, 'confidence': score, 'bbox': bbox})
        return out
    if isinstance(result, dict):
        texts = result.get('rec_texts', [])
        scores = result.get('rec_scores', [])
        polys = result.get('rec_polys', [])
        for i, (text, score) in enumerate(zip(texts, scores)):
            bbox = polys[i] if i < len(polys) else []
            out.append({'text': text, 'confidence': score, 'bbox': bbox})
        return out
    if isinstance(result, (list, tuple)):
        # 2.x 格式: [box, (text, score)]
        if (len(result) >= 2 and isinstance(result[0], (list, tuple))
                and isinstance(result[1], (list, tuple))
                and len(result[1]) >= 2):
            out.append({'text': str(result[1][0]),
                        'confidence': float(result[1][1]),
                        'bbox': result[0]})
        return out
    return out


# ---------- 图像预处理 ----------
def _preprocess_image(pil_image, preprocess):
    """对 PIL Image 做预处理，返回处理后的 Image。"""
    if preprocess == 'grayscale':
        return pil_image.convert('L').convert('RGB')
    if preprocess == 'otsu':
        import numpy as np
        gray = pil_image.convert('L')
        arr = np.array(gray)
        # 大津法二值化
        hist, _ = np.histogram(arr.ravel(), 256, [0, 256])
        total = arr.size
        sum_all = np.dot(np.arange(256), hist)
        w_f, s_f = 0.0, 0.0
        max_var, threshold = 0.0, 0
        for t in range(256):
            w_f += hist[t]
            if w_f == 0 or w_f == total:
                continue
            w_b = total - w_f
            s_f += t * hist[t]
            m_f = s_f / w_f
            m_b = (sum_all - s_f) / w_b
            var_between = w_f * w_b * (m_f - m_b) ** 2
            if var_between > max_var:
                max_var, threshold = var_between, t
        binary = (arr > threshold).astype('uint8') * 255
        from PIL import Image
        return Image.fromarray(binary, mode='L').convert('RGB')
    if preprocess == 'stretch':
        # 直方图拉伸：处理显微镜图像偏暗/偏亮的问题
        # 对 RGB 各通道分别拉伸（保留颜色差异）
        import numpy as np
        arr = np.array(pil_image)
        lo, hi = arr.min(), arr.max()
        if hi - lo < 5:
            return pil_image  # 对比度太小，跳过
        stretched = ((arr.astype('float32') - lo) / (hi - lo) * 255)
        stretched = stretched.clip(0, 255).astype('uint8')
        from PIL import Image
        return Image.fromarray(stretched)
    if preprocess == 'stretch_invert':
        # 直方图拉伸 + 反转（适合暗底亮字的显微镜图像）
        import numpy as np
        arr = np.array(pil_image)
        lo, hi = arr.min(), arr.max()
        if hi - lo < 5:
            return pil_image
        stretched = ((arr.astype('float32') - lo) / (hi - lo) * 255)
        stretched = 255 - stretched.clip(0, 255).astype('uint8')
        from PIL import Image
        return Image.fromarray(stretched)
    return pil_image


def _crop_roi(pil_image, roi):
    """按 (x, y, w, h) 裁剪图像。roi 为 None 返回原图。"""
    if roi is None:
        return pil_image
    x, y, w, h = roi
    # clamp 到图像边界
    iw, ih = pil_image.size
    x = max(0, min(x, iw - 1))
    y = max(0, min(y, ih - 1))
    w = max(1, min(w, iw - x))
    h = max(1, min(h, ih - y))
    return pil_image.crop((x, y, x + w, y + h))


# ---------- OCR ----------
def ocr_image(image_path, roi=None, preprocess='none', lang='ch'):
    """对单张图像做 OCR，返回检测到的文本列表。

    返回: [{'text': str, 'confidence': float, 'bbox': [x1,y1,x2,y2]}, ...]
    """
    ocr = _get_ocr(lang)
    from PIL import Image
    img = Image.open(image_path)
    if img.mode != 'RGB':
        img = img.convert('RGB')
    if roi is not None:
        img = _crop_roi(img, roi)
    img = _preprocess_image(img, preprocess)
    # PaddleOCR 接受 ndarray
    import numpy as np
    arr = np.array(img)
    results = _run_ocr(ocr, arr)
    if not results:
        return []
    out = []
    for result in results:
        out.extend(_normalize_ocr_result(result))
    return out


# ---------- 数值提取 ----------
_RE_VALUE = re.compile(r'[=:：]\s*(\d+(?:\.\d+)?)')
_RE_NUMBER = re.compile(r'\d+(?:\.\d+)?')


def _parse_number(text):
    """从字符串中提取数值：优先取 '=155.76' 等分隔符后的值，
    避免把标签里的数字（如 No.2）当成量测值。"""
    m = _RE_VALUE.search(text)
    if m:
        s = m.group(1)
    else:
        m = _RE_NUMBER.search(text)
        if not m:
            return None
        s = m.group(0)
    # 去掉末尾多余的点
    s = s.rstrip('.')
    if not s:
        return None
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        return None


def _build_label_pattern(label):
    """把标签（如 'No.2'）编译为正则，匹配 'No.2=155.76' 等形式。

    支持分隔符: = : ： 及可选的空格。
    """
    escaped = re.escape(label)
    return re.compile(
        rf'{escaped}\s*[=:：]\s*([\d.]+)',
        re.IGNORECASE
    )


def extract_values(ocr_results, mode='first_number', labels=None, expr=''):
    """从 OCR 结果中提取量测值。

    Args:
        ocr_results: ocr_image() 的返回列表
        mode: 'first_number' | 'labeled' | 'all_numbers' | 'custom'
        labels: labeled 模式下的标签列表，如 ['No.2', 'No.3']
        expr: custom 模式下的表达式（x = 所有文本用空格拼接）

    Returns:
        - first_number:  {'value': 155.76} 或未找到时 {'value': None}
        - labeled:       {'No.2': 155.76, 'No.3': 2.30, ...}
        - all_numbers:   {'values': [155.76, 56.91, ...]}
        - custom:        {'value': <表达式结果>}
    """
    all_text = ' '.join(r['text'] for r in ocr_results)

    if mode == 'first_number':
        return {'value': _parse_number(all_text)}

    if mode == 'labeled':
        if not labels:
            return {}
        result = {}
        for label in labels:
            pat = _build_label_pattern(label)
            match = pat.search(all_text)
            if match:
                result[label] = _parse_number(match.group(1))
            else:
                result[label] = None
        return result

    if mode == 'all_numbers':
        nums = []
        for r in ocr_results:
            v = _parse_number(r['text'])
            if v is not None:
                nums.append(v)
        return {'values': nums}

    if mode == 'custom':
        if not expr.strip():
            return {'value': all_text}
        from safe_eval import _safe_eval_transform
        ok, result = _safe_eval_transform(expr.strip(), all_text)
        return {'value': result if ok else None}

    return {}


def ocr_batch(image_paths, roi=None, preprocess='none', lang='ch',
              mode='first_number', labels=None, expr='',
              progress_callback=None):
    """批量 OCR，返回每张图的结果列表。

    progress_callback(current, total) 在每个图像完成后调用。
    """
    results = []
    total = len(image_paths)
    for i, path in enumerate(image_paths):
        try:
            ocr_results = ocr_image(path, roi=roi, preprocess=preprocess,
                                     lang=lang)
            extracted = extract_values(ocr_results, mode=mode,
                                        labels=labels, expr=expr)
            extracted['_image'] = os.path.basename(path)
            extracted['_path'] = path
            extracted['_ocr_texts'] = [r['text'] for r in ocr_results]
            results.append(extracted)
        except Exception as e:
            logger.warning(f"OCR 失败 {path}: {e}")
            results.append({
                '_image': os.path.basename(path),
                '_path': path,
                '_error': str(e),
                '_ocr_texts': [],
            })
        if progress_callback:
            progress_callback(i + 1, total)
    return results


# ---------- 带独立 ROI 的批量 OCR（多值模式） ----------
def ocr_batch_with_rois(image_paths, roi_configs=None, preprocess='none',
                         lang='ch', mode='first_number', labels=None, expr='',
                         progress_callback=None):
    """批量 OCR，每张图可用不同的 ROI 配置。

    roi_configs: 与 image_paths 等长的列表，每项为 roi tuple 或 None，
                 或包含 {label: roi} 的 dict（按标签不同ROI）。
    """
    results = []
    total = len(image_paths)
    for i, path in enumerate(image_paths):
        try:
            roi_cfg = roi_configs[i] if roi_configs and i < len(roi_configs) else None
            if isinstance(roi_cfg, dict) and mode == 'labeled' and labels:
                # 按标签独立 ROI：对每个标签分别 OCR
                all_ocr = []
                for label in labels:
                    lr = roi_cfg.get(label, roi_cfg.get('*'))
                    label_results = ocr_image(path, roi=lr,
                                              preprocess=preprocess, lang=lang)
                    all_ocr.extend(label_results)
                extracted = extract_values(all_ocr, mode=mode,
                                            labels=labels, expr=expr)
            else:
                roi = roi_cfg if not isinstance(roi_cfg, dict) else roi_cfg.get('*')
                ocr_results = ocr_image(path, roi=roi,
                                         preprocess=preprocess, lang=lang)
                extracted = extract_values(ocr_results, mode=mode,
                                            labels=labels, expr=expr)
            extracted['_image'] = os.path.basename(path)
            extracted['_path'] = path
            results.append(extracted)
        except Exception as e:
            logger.warning(f"OCR 失败 {path}: {e}")
            results.append({
                '_image': os.path.basename(path),
                '_path': path,
                '_error': str(e),
            })
        if progress_callback:
            progress_callback(i + 1, total)
    return results
