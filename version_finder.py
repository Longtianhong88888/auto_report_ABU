"""版本号提取工具。

从用户指定的档案（PDF / Excel）中提取 Summary 的四个版本号：
- Process control rev.  -> ERS 档案内部 "Process Control Specification" 区块的 Cx.y
- ERS rev. / VSR rev.   -> 文件名 RevN 优先，其次正文中的最大 RevN
- MCO rev.              -> 文件名末尾 "-NN"（如 056-25088-02 -> 2）
"""
import os
import re
import zlib
from functools import lru_cache

import openpyxl


REV_FILE = re.compile(r'_?Rev\.?\s*(\d+)', re.I)
MCO_FILE = re.compile(r'-(\d{1,4})\.(?:pdf|xlsx)$', re.I)
PC_TABLE = re.compile(r'TABLE\s*:?\s*PROCESSCONTROL\s*REV\s*(\d+(?:\.\d+)?)')
REV_TEXT = re.compile(r'REV\.?\s*(\d+)')


def _extract_pdf_streams(path):
    """提取 PDF 中的 FlateDecode 文本流（跳过图片流，加速且避免乱码）"""
    with open(path, 'rb') as f:
        data = f.read()
    streams = []
    for m in re.finditer(rb'<<(.*?)>>\s*stream\r?\n(.*?)endstream', data, re.S):
        if b'/Subtype' in m.group(1) and b'/Image' in m.group(1):
            continue
        raw = m.group(2)
        try:
            streams.append(zlib.decompress(raw))
        except Exception:
            continue
    return streams


def _pdf_strings(path):
    """把 PDF 文本流中的 (…)Tj / [(…)TJ 文本提取出来"""
    out = []
    pat = re.compile(r'\(((?:[^()\\]|\\.)*)\)\s*(?:Tj|TJ)?')
    for stream in _extract_pdf_streams(path):
        try:
            text = stream.decode('latin-1')
        except Exception:
            continue
        for m in pat.finditer(text):
            s = re.sub(r'\\([()\\])', r'\1', m.group(1))
            if s.strip():
                out.append(s)
    return out


def _xlsx_strings(path):
    """把 Excel 所有单元格文本拼起来，便于统一做正则匹配"""
    out = []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        out.append(cell.value)
    finally:
        wb.close()
    return out


def _normalize(text):
    """去掉分隔符/空白，便于匹配被拆分的 PDF 文本（如 'C | 7 | .0'）"""
    return re.sub(r'[^A-Za-z0-9.]', '', text).upper()


@lru_cache(maxsize=32)
def _raw_text(path):
    """档案原始文本（PDF 文本片段 / Excel 单元格文本）"""
    if path.lower().endswith('.pdf'):
        return '\n'.join(_pdf_strings(path))
    if path.lower().endswith(('.xlsx', '.xlsm')):
        return '\n'.join(_xlsx_strings(path))
    return ''


@lru_cache(maxsize=32)
def _normalized_text(path):
    return _normalize(_raw_text(path))


def detect_process_control(path):
    """Process control rev.：匹配 'Table : Process Control Rev x.xx'。
    文档中该表头会在每页重复出现，多个版本取最大值（如 BUF ERS -> 3.52）。"""
    versions = [float(x) for x in PC_TABLE.findall(_normalized_text(path))]
    if versions:
        value = max(versions)
        return int(value) if value == int(value) else round(value, 2)
    return None


def detect_rev_number(path):
    """ERS / VSR rev.：文件名 RevN 优先，其次正文中最大的 RevN"""
    m = REV_FILE.search(os.path.basename(path))
    if m:
        return int(m.group(1))
    nums = [int(x) for x in REV_TEXT.findall(_normalized_text(path))]
    if nums:
        return max(nums)
    return None


def detect_mco_rev(path):
    """MCO rev.：文件名末尾段 '-NN'（如 056-25088-02 -> 2），
    其次文件名 RevN，最后正文 RevN"""
    m = MCO_FILE.search(os.path.basename(path))
    if m:
        return int(m.group(1))
    m = REV_FILE.search(os.path.basename(path))
    if m:
        return int(m.group(1))
    nums = [int(x) for x in REV_TEXT.findall(_normalized_text(path))]
    if nums:
        return max(nums)
    return None


def detect_version(field, path):
    """按字段名分派提取"""
    if field == 'process_control':
        return detect_process_control(path)
    if field == 'ers':
        return detect_rev_number(path)
    if field == 'vsr':
        return detect_rev_number(path)
    if field == 'mco':
        return detect_mco_rev(path)
    return None


def suggest_files(base_dir, keyword='BUF'):
    """在目录（含同名子文件夹）中推荐 ERS/VSR/MCO 档案。
    keyword='BUF' 时优先匹配文件名含 BUF 的（其他专案通用）；
    keyword='CLO' 时优先匹配 CLO 系。"""
    if not base_dir or not os.path.isdir(base_dir):
        return {}
    suggestions = {}
    for key in ('ers', 'vsr', 'mco'):
        candidates = []
        try:
            entries = sorted(os.listdir(base_dir))
        except OSError:
            continue
        for entry in entries:
            full = os.path.join(base_dir, entry)
            if os.path.isdir(full) and entry.lower() == key:
                try:
                    candidates.extend(
                        os.path.join(full, f) for f in sorted(os.listdir(full))
                        if f.lower().endswith(('.pdf', '.xlsx')))
                except OSError:
                    pass
            elif (os.path.isfile(full)
                  and key in entry.lower()
                  and entry.lower().endswith(('.pdf', '.xlsx'))):
                candidates.append(full)
        if candidates:
            preferred = [p for p in candidates
                         if keyword in os.path.basename(p).upper()]
            suggestions[key] = (preferred or candidates)[0]
    return suggestions
