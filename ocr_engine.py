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
import multiprocessing
from functools import lru_cache
from queue import Empty

from constants import OCR_MAX_IMAGE_SIDE

# Windows/Linux x86_64 上 PaddlePaddle 3.3.0+ 的 PIR→oneDNN(MKLDNN) 回归
# 会使 PP-OCRv6 CPU 推理直接崩溃（ConvertPirAttribute2RuntimeAttribute）。
# 关键：paddlex 在“模块导入时”读取该开关（paddlex.utils.flags），
# 必须放在任何 paddleocr/paddlex import 之前；若等 _get_ocr() 里再设，
# ocr_available() 已先 import 过 paddleocr，开关会失效。
os.environ.setdefault('PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT', 'false')
# PyInstaller 打包后 OpenCV 与 Paddle 各带一份 OpenMP 运行库，冲突会在
# predict 时静默崩溃（Windows 0xC0000005），放开重复库加载限制。
os.environ.setdefault('KMP_DUPLICATE_LIB_OK', 'TRUE')

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
        # 防御：即使模块被其它入口先导入，这里再次确保开关为 false
        os.environ.setdefault('PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT', 'false')
        os.environ.setdefault('KMP_DUPLICATE_LIB_OK', 'TRUE')
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
            enable_mkldnn=False,
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
    """检查 PaddleOCR 是否可用（不触发模型下载，也不导入 paddle 本体）。

    识别在独立子进程中执行，父进程不应加载 paddle 的原生库；
    这里只做模块定位，避免 DLL 加载失败/崩溃影响主程序。
    """
    try:
        import importlib.util
        return importlib.util.find_spec('paddleocr') is not None
    except (ImportError, ValueError, AttributeError):
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


def _limit_image_size(pil_image, max_side=OCR_MAX_IMAGE_SIDE):
    """识别前把图像最长边等比压缩到 max_side 以内。

    高分辨率显微镜大图（几千万像素）直接转 numpy 会占用大量内存，
    而 PaddleOCR 检测阶段内部本来就会缩放到约 736~960 边长再推理，
    所以先压缩到 2048 以内不影响识别质量，还能大幅降低内存占用。
    ROI 裁剪之后调用，坐标语义不变。"""
    w, h = pil_image.size
    longest = max(w, h)
    if longest <= max_side:
        return pil_image
    scale = max_side / float(longest)
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))
    from PIL import Image
    return pil_image.resize((new_w, new_h), Image.LANCZOS)


# ---------- OCR ----------
def ocr_image(image_path, roi=None, preprocess='none', lang='ch'):
    """对单张图像做 OCR，返回检测到的文本列表。

    返回: [{'text': str, 'confidence': float, 'bbox': [x1,y1,x2,y2]}, ...]
    """
    ocr = _get_ocr(lang)
    from PIL import Image
    # 高分辨率切片大图常超过 Pillow 默认像素上限，关闭 DecompressionBomb 限制
    Image.MAX_IMAGE_PIXELS = None
    img = Image.open(image_path)
    if img.mode != 'RGB':
        img = img.convert('RGB')
    if roi is not None:
        img = _crop_roi(img, roi)
    # 超大图先压缩再预处理/转数组，避免识别时内存占用过高
    img = _limit_image_size(img)
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
_RE_AFTER_BRACKET = re.compile(r'\](\d+(?:\.\d+)?)')  # [1]470.41 格式
_RE_NUMBER = re.compile(r'\d+(?:\.\d+)?')


def _parse_number(text):
    """从字符串中提取数值。

    优先级：
    1. '=155.76' / ':155.76' 等显式分隔符后的值
    2. '[1]470.41' 右括号后的值（避免把标签编号当测量值）
    3. 所有数值中取最大的（测量值通常 > 标签编号）
    """
    m = _RE_VALUE.search(text)
    if m:
        s = m.group(1)
    else:
        m = _RE_AFTER_BRACKET.search(text)
        if m:
            s = m.group(1)
        else:
            numbers = _RE_NUMBER.findall(text)
            if not numbers:
                return None
            s = max(numbers, key=lambda n: float(n))
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
    """把标签（如 'No.2', '[1]'）编译为正则，匹配 'No.2=155.76'、'[1]470.41μm' 等形式。

    支持分隔符: = : ： 空格，或标签与数值直接相邻（无分隔符）。
    """
    escaped = re.escape(label)
    return re.compile(
        rf'{escaped}\b\s*[=:：\s]*([\d.]+)',
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


# ---------- OCR 子进程隔离 ----------
# PaddleOCR/Paddle 在 Windows CPU 上存在已知原生崩溃（PIR→oneDNN 回归、
# OpenMP 重复库冲突等，0xC0000005），Python 异常钩子无法拦截。
# 因此把所有识别放到独立子进程执行：子进程崩溃只影响它自己，
# 父进程检测到异常退出/超时后返回错误结果，主程序与已做的映射配置不受影响。

OCR_SUBPROCESS_INIT_TIMEOUT = 300    # 首次加载模型（冷启动）等待上限（秒）
OCR_SUBPROCESS_IMAGE_TIMEOUT = 120   # 单张图识别等待上限（秒）


class OCRSubprocessError(RuntimeError):
    """OCR 子进程失败/崩溃/超时（已隔离，主程序可继续运行）"""


def _iter_ocr_batch(image_paths, roi=None, preprocess='none', lang='ch',
                    mode='first_number', labels=None, expr=''):
    """子进程内逐张 OCR（生成器，逐张回传结果供父进程更新进度）"""
    for path in image_paths:
        try:
            ocr_results = ocr_image(path, roi=roi, preprocess=preprocess,
                                    lang=lang)
            extracted = extract_values(ocr_results, mode=mode,
                                       labels=labels, expr=expr)
            extracted['_image'] = os.path.basename(path)
            extracted['_path'] = path
            extracted['_ocr_texts'] = [r['text'] for r in ocr_results]
            yield extracted
        except Exception as e:
            logger.warning(f"OCR 失败 {path}: {e}")
            yield {
                '_image': os.path.basename(path),
                '_path': path,
                '_error': str(e),
                '_ocr_texts': [],
            }


def _iter_ocr_batch_with_rois(image_paths, roi_configs=None, preprocess='none',
                              lang='ch', mode='first_number', labels=None,
                              expr=''):
    """子进程内批量 OCR（每张图可独立 ROI），逐张回传结果"""
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
            yield extracted
        except Exception as e:
            logger.warning(f"OCR 失败 {path}: {e}")
            yield {
                '_image': os.path.basename(path),
                '_path': path,
                '_error': str(e),
            }


def _ocr_worker_main(task_queue, result_queue):
    """OCR 子进程入口：从队列取任务，逐张回传结果。

    该函数在独立进程中运行，import 的是被 spawn 的 __mp_main__
    （即 main.py 顶层，只做导入与环境变量设置，不会启动主界面）。
    PaddleOCR 原生崩溃时整个子进程退出，由父进程检测兜底。
    每次调用只处理一个批量任务，完成后立即退出，方便父进程回收。
    """
    try:
        task = task_queue.get()
    except (EOFError, OSError):
        return
    if task is None:
        return
    func_name, kwargs = task
    try:
        if func_name == 'batch':
            iterator = _iter_ocr_batch(**kwargs)
        elif func_name == 'batch_with_rois':
            iterator = _iter_ocr_batch_with_rois(**kwargs)
        else:
            raise ValueError(f"未知 OCR 任务: {func_name!r}")
        for result in iterator:
            result_queue.put({'type': 'result', 'result': result})
    except Exception as e:
        # 任务级致命错误（如模型初始化失败），回传后退出本进程
        result_queue.put({'type': 'fatal_error',
                          'error': f'{type(e).__name__}: {e}'})


def _run_ocr_subprocess(func_name, image_paths, kwargs, total,
                        progress_callback=None):
    """启动 OCR 子进程执行批量识别，父进程逐张收集结果。

    子进程正常返回时逐张回传；崩溃/超时时抛 OCRSubprocessError，
    由调用方（预览 worker / 输出流程）转成可读错误，主程序不退出。
    """
    if total <= 0:
        return []
    ctx = multiprocessing.get_context('spawn')
    task_queue = ctx.Queue()
    result_queue = ctx.Queue()
    proc = ctx.Process(target=_ocr_worker_main,
                       args=(task_queue, result_queue), daemon=True)
    proc.start()
    results = []
    try:
        task_queue.put((func_name, {**kwargs, 'image_paths': image_paths}))
        # 首条结果放宽（含模型加载），之后按单张图放宽
        timeout = OCR_SUBPROCESS_INIT_TIMEOUT
        while len(results) < total:
            try:
                msg = result_queue.get(timeout=timeout)
            except Empty:
                if proc.is_alive():
                    proc.terminate()
                    proc.join(10)
                    raise OCRSubprocessError(
                        f"OCR 子进程超时（{timeout}s 无响应），已终止")
                raise OCRSubprocessError(
                    f"OCR 子进程提前退出（exitcode={proc.exitcode}），"
                    "识别已隔离，主程序不受影响")
            if msg.get('type') == 'fatal_error':
                raise OCRSubprocessError(f"OCR 子进程错误：{msg.get('error')}")
            results.append(msg.get('result'))
            if progress_callback:
                progress_callback(len(results), total)
            timeout = OCR_SUBPROCESS_IMAGE_TIMEOUT
        proc.join(10)
    finally:
        if proc.is_alive():
            proc.terminate()
            proc.join(10)
        try:
            task_queue.close()
            result_queue.close()
            task_queue.join_thread()
            result_queue.join_thread()
        except Exception:
            pass
    return results


def ocr_batch(image_paths, roi=None, preprocess='none', lang='ch',
              mode='first_number', labels=None, expr='',
              progress_callback=None):
    """批量 OCR（子进程隔离），返回每张图的结果列表。

    progress_callback(current, total) 在每个图像完成后调用。
    PaddleOCR 原生崩溃不会影响主程序，而是以 OCRSubprocessError 抛出。
    """
    total = len(image_paths)
    kwargs = dict(roi=roi, preprocess=preprocess, lang=lang,
                  mode=mode, labels=labels, expr=expr)
    return _run_ocr_subprocess('batch', image_paths, kwargs, total,
                               progress_callback)


# ---------- 带独立 ROI 的批量 OCR（多值模式） ----------
def ocr_batch_with_rois(image_paths, roi_configs=None, preprocess='none',
                        lang='ch', mode='first_number', labels=None, expr='',
                        progress_callback=None):
    """批量 OCR（子进程隔离），每张图可用不同的 ROI 配置。

    roi_configs: 与 image_paths 等长的列表，每项为 roi tuple 或 None，
                 或包含 {label: roi} 的 dict（按标签不同ROI）。
    """
    total = len(image_paths)
    kwargs = dict(roi_configs=roi_configs, preprocess=preprocess, lang=lang,
                  mode=mode, labels=labels, expr=expr)
    return _run_ocr_subprocess('batch_with_rois', image_paths, kwargs, total,
                               progress_callback)
