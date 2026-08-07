"""公共工具：列宽解析、表格统一尺寸、映射配置标准化。"""
from openpyxl.utils import get_column_letter

from constants import (
    COL_WIDTH_PX_PER_CHAR, ROW_HEIGHT_PX_PER_PT,
    DEFAULT_COL_WIDTH_CHARS, DEFAULT_ROW_HEIGHT_PTS,
)


def column_width_chars(ws, col_idx):
    """解析某列宽度（字符数），兼容 openpyxl 不展开的区间列定义。
    Excel 常写成 <col min="3" max="13" width="33"/>，openpyxl 只保留
    首列索引，其余列查不到，需要遍历全部维度按 min/max 区间匹配。
    未定义列退回 sheet 默认列宽。"""
    letter = get_column_letter(col_idx)
    if letter in ws.column_dimensions:
        dim = ws.column_dimensions[letter]
        if dim.width:
            return dim.width
    for dim in ws.column_dimensions.values():
        lo = getattr(dim, 'min', None) or 0
        hi = getattr(dim, 'max', None) or lo
        if lo <= col_idx <= hi and dim.width:
            return dim.width
    return ws.sheet_format.defaultColWidth or DEFAULT_COL_WIDTH_CHARS


def apply_uniform_sizes(table, ws, max_cols, max_rows, zoom=None):
    """统一设置表格列宽与行高（未显式定义的行高回落 sheet 默认值）。
    zoom 传入 TableZoomMixin 实例时，同步记录基准尺寸并应用当前缩放。"""
    for col in range(1, max_cols + 1):
        width = int(column_width_chars(ws, col) * COL_WIDTH_PX_PER_CHAR)
        if zoom is not None:
            zoom.zoom_size(col - 1, width, None, None)
        else:
            table.setColumnWidth(col - 1, width)
    for row in range(1, max_rows + 1):
        if row in ws.row_dimensions and ws.row_dimensions[row].height:
            height = int(ws.row_dimensions[row].height * ROW_HEIGHT_PX_PER_PT)
        else:
            default_pts = ws.sheet_format.defaultRowHeight or DEFAULT_ROW_HEIGHT_PTS
            height = int(default_pts * ROW_HEIGHT_PX_PER_PT)
        if zoom is not None:
            zoom.zoom_size(None, None, row - 1, height)
        else:
            table.setRowHeight(row - 1, height)


def normalize_mappings(mappings):
    """配置导入时的映射标准化：
    - JMP 的 header_cols 统一为字符串列表（兼容旧配置存字符串）；
    - 归档/数据的区间字段统一为元组。"""
    result = []
    for m in mappings:
        if not isinstance(m, dict):
            continue
        m = dict(m)
        mtype = m.get('type')
        if mtype == 'jmp':
            headers = m.get('header_cols')
            if isinstance(headers, str):
                m['header_cols'] = [headers]
            elif isinstance(headers, (list, tuple)):
                m['header_cols'] = [str(h) for h in headers]
            else:
                m['header_cols'] = []
        for field in ('block_range', 'target_range'):
            if field in m and isinstance(m[field], list):
                m[field] = tuple(m[field])
        result.append(m)
    return result
