import io
import os
import re
from copy import copy
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Border, Side
from PIL import Image as PILImage
from constants import COL_WIDTH_PX_PER_CHAR, ROW_HEIGHT_PX_PER_PT
from safe_eval import _safe_eval_transform

class MappingOperations:
    """数据/图片/归档/JMP 映射执行逻辑（供 MainWindow 混入）"""
    @staticmethod
    def _archive_blocks_overlap(block1, block2):
        """两个归档区块是否重叠：行区间与列区间都相交才算重叠"""
        r1_min, c1_min, r1_max, c1_max = block1
        r2_min, c2_min, r2_max, c2_max = block2
        return (r1_min <= r2_max and r2_min <= r1_max
                and c1_min <= c2_max and c2_min <= c1_max)
    @staticmethod
    def _block_range_str(block_range):
        min_row, min_col, max_row, max_col = block_range
        return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    @staticmethod
    def _mapping_key(m):
        """映射去重键：同类型、同目标、同位置视为同一条映射"""
        t = m.get('type')
        if t == 'archive_shift_right':
            return ('archive', m.get('target_sheet'), tuple(m.get('block_range')))
        if t == 'data':
            return ('data', m.get('target_sheet'), tuple(m.get('target_range')),
                    m.get('source_sheet'), m.get('source_range'), m.get('transform'))
        if t == 'image':
            return ('image', m.get('target_sheet'), m.get('anchor_cell'))
        if t == 'jmp':
            return ('jmp', m.get('target_sheet'), m.get('anchor_cell'))
        return (t, id(m))
    def _dedupe_mappings(self, mappings):
        """按映射去重键过滤重复项，保留首次出现"""
        seen = set()
        result = []
        for m in mappings:
            key = self._mapping_key(m)
            if key in seen:
                continue
            seen.add(key)
            result.append(m)
        return result
    @staticmethod
    def _extend_block_top_merges(ws, min_row, min_col, max_row, max_col):
        """把覆盖区块首行（或恰好位于首行上方一行，如组表头）的合并单元格向右扩一格。
        与已有合并冲突时跳过，避免合并范围互相覆盖。"""
        targets = []
        for mr in ws.merged_cells.ranges:
            m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(str(mr))
            # 行：覆盖首行 或 位于首行上方一行；列：与区块列区间相交
            if not (m_min_row <= min_row <= m_max_row or m_min_row == min_row - 1):
                continue
            if not (m_min_col <= max_col and m_max_col >= min_col):
                continue
            targets.append((str(mr), m_min_row, m_min_col, m_max_row, m_max_col))
        for mr_str, m_min_row, m_min_col, m_max_row, m_max_col in targets:
            # 目标列 m_max_col+1 若已被其他合并占用则跳过
            conflict = False
            for other in ws.merged_cells.ranges:
                if str(other) == mr_str:
                    continue
                o_min_col, o_min_row, o_max_col, o_max_row = range_boundaries(str(other))
                if (m_min_row <= o_max_row and o_min_row <= m_max_row
                        and o_min_col <= m_max_col + 1 <= o_max_col):
                    conflict = True
                    break
            if conflict:
                continue
            ws.unmerge_cells(mr_str)
            ws.merge_cells(start_row=m_min_row, start_column=m_min_col,
                           end_row=m_max_row, end_column=m_max_col + 1)
    @staticmethod
    def _apply_block_borders(ws, block_range):
        """归档后：先对数据块全部单元格加细框线，再对外边界加粗框线"""
        min_row, min_col, max_row, max_col = block_range
        thin = Side(style='thin', color='000000')
        thick = Side(style='thick', color='000000')
        # 1) 全部框线（细线）
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).border = Border(
                    left=thin, right=thin, top=thin, bottom=thin)
        # 2) 粗外框（覆盖外边界）
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                b = cell.border
                cell.border = Border(
                    left=thick if c == min_col else b.left,
                    right=thick if c == max_col else b.right,
                    top=thick if r == min_row else b.top,
                    bottom=thick if r == max_row else b.bottom,
                )
    def execute_mapping(self, ws, mapping, data_ws, data_src_ws=None, jmp_src_ws=None):
        if mapping['type'] == 'data':
            self.apply_data_mapping(ws, mapping, data_src_ws)
        elif mapping['type'] == 'image':
            self.apply_image_mapping(ws, mapping)
        elif mapping['type'] == 'archive_shift_right':
            self.apply_archive_shift_right(ws, mapping, data_ws, data_src_ws)
        elif mapping['type'] == 'jmp':
            self.apply_jmp_mapping(ws, mapping, jmp_src_ws)
    def _resolve_data_src_ws(self, mapping, data_source_wb):
        """按映射的 source_sheet 从数据源文件解析工作表（优先 data_only 版本）"""
        src_sheet = mapping.get('source_sheet')
        if data_source_wb and src_sheet and src_sheet in data_source_wb.sheetnames:
            return data_source_wb[src_sheet]
        return None
    @staticmethod
    def _excel_error(value):
        """识别 Excel 错误值字符串；非错误值返回 None"""
        if isinstance(value, str) and value in (
            '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!'
        ):
            return value
        return None
    def _warn_error_value(self, value, src_addr, dst_addr):
        """若填充值来自数据源且是 Excel 错误值，记录警告供输出时提示"""
        err = self._excel_error(value)
        if err:
            self._fill_warnings.append(f"{dst_addr} <- {src_addr}：数据源错误值 {err}")
        # ---------- 数据转换 ----------
    def _apply_transform(self, value, mapping):
        trans = mapping.get('transform', 'none')
        if trans == 'none' or value is None:
            return value
        try:
            if trans == 'div1000':
                return float(value) / 1000.0
            elif trans == 'strip_letters':
                if isinstance(value, str):
                    return re.sub(r'[a-zA-Z]+$', '', value)
                return value
            elif trans == 'custom':
                expr = mapping.get('transform_expr', '')
                if expr:
                    ok, result = _safe_eval_transform(expr, value)
                    if ok:
                        return result
                    msg = f"自定义表达式计算失败，已保留原值：{expr!r}"
                    if msg not in self._fill_warnings:
                        self._fill_warnings.append(msg)
        except:
            pass
        return value
    def apply_data_mapping(self, ws, mapping, data_src_ws=None):
        t_min_row, t_min_col, t_max_row, t_max_col = mapping['target_range']
        src_ws = data_src_ws if data_src_ws else self.source_wb[mapping['source_sheet']]
        s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(mapping['source_range'])
        for i in range(t_max_row - t_min_row + 1):
            for j in range(t_max_col - t_min_col + 1):
                src_row = s_min_row + i
                src_col = s_min_col + j
                if src_row <= s_max_row and src_col <= s_max_col:
                    val = src_ws.cell(row=src_row, column=src_col).value
                    val = self._apply_transform(val, mapping)
                    dst_row = t_min_row + i
                    dst_col = t_min_col + j
                    self._warn_error_value(val,
                        f"{src_ws.title}!{get_column_letter(src_col)}{src_row}",
                        f"{ws.title}!{get_column_letter(dst_col)}{dst_row}")
                    ws.cell(row=dst_row, column=dst_col).value = val
    def remove_images_at_anchor(self, ws, anchor):
        to_remove = []
        for img in ws._images:
            if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                r = img.anchor._from.row + 1
                c = img.anchor._from.col + 1
                if f"{get_column_letter(c)}{r}" == anchor:
                    to_remove.append(img)
        for img in to_remove:
            ws._images.remove(img)
    def _process_image_data(self, img_bytes, rotation, target_w, target_h):
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        if rotation != 0:
            pil_img = pil_img.rotate(-rotation, expand=True)
        pil_img = pil_img.resize((target_w, target_h), PILImage.LANCZOS)
        out_stream = io.BytesIO()
        pil_img.save(out_stream, format='PNG')
        out_stream.seek(0)
        return out_stream
    def apply_image_mapping(self, ws, mapping):
        try:
            anchor = mapping['anchor_cell']
            col_width_chars = mapping.get('col_width_chars', 8.0)
            row_height_pts = mapping.get('row_height_pts', 15.0)
            rotation = mapping.get('rotation', 0.0)
            w_scale = mapping.get('width_scale', 1.0)
            h_scale = mapping.get('height_scale', 1.0)
            target_width = int(col_width_chars * COL_WIDTH_PX_PER_CHAR * w_scale)
            target_height = int(row_height_pts * ROW_HEIGHT_PX_PER_PT * h_scale)
            if 'image_path' in mapping:
                if not os.path.exists(mapping['image_path']):
                    raise FileNotFoundError(f"图片文件不存在: {mapping['image_path']}")
                with open(mapping['image_path'], 'rb') as f:
                    img_bytes = f.read()
                processed_stream = self._process_image_data(img_bytes, rotation, target_width, target_height)
                self._image_streams.append(processed_stream)
                self.remove_images_at_anchor(ws, anchor)
                new_img = OpenpyxlImage(processed_stream)
                new_img._saved_bytes = processed_stream.getvalue()
                new_img.anchor = anchor
                ws.add_image(new_img)
                return
            # 配置只记录映射路径：优先按“源Sheet + 锚点位置”从当前数据源
            # 重新读取图片内容，旧配置兼容 image_bytes / image_ref
            img_bytes = None
            if mapping.get('image_src_sheet') and mapping.get('image_src_pos'):
                img_bytes = self._get_image_bytes_by_pos(
                    mapping['image_src_sheet'], mapping['image_src_pos'])
            elif mapping.get('image_bytes'):
                img_bytes = mapping['image_bytes']
            elif 'image_ref' in mapping and self.source_wb:
                img_bytes = self._get_internal_image_bytes(mapping['image_ref'])
            if img_bytes is None:
                raise RuntimeError("映射中缺少图片数据（请先打开 IPQC 数据源或重新选择图片）")
            processed_stream = self._process_image_data(img_bytes, rotation, target_width, target_height)
            self._image_streams.append(processed_stream)
            self.remove_images_at_anchor(ws, anchor)
            new_img = OpenpyxlImage(processed_stream)
            new_img._saved_bytes = processed_stream.getvalue()
            new_img.anchor = anchor
            ws.add_image(new_img)
        except Exception as e:
            self._fill_warnings.append(f"图片写入失败 {mapping.get('anchor_cell', '?')}：{e}")
    def _get_internal_image_bytes(self, image_ref):
        """按 (sheet, 序号) 从缓存中取图片字节（避免读取已被 openpyxl 关闭的 ref）"""
        sheet_name, idx = image_ref
        for cached in self.cached_images:
            if cached[0] == sheet_name and cached[1] == idx:
                return cached[3]
        raise RuntimeError(f"未找到图片 {sheet_name}[{idx}]")

    def _get_image_bytes_by_pos(self, sheet_name, pos):
        """按 (sheet, 锚点位置) 从缓存中取图片字节。
        配置只记录来源路径，内容随当前数据源更新，避免旧映射图片缺失。"""
        for cached in self.cached_images:
            if cached[0] == sheet_name and cached[2] == pos:
                return cached[3]
        raise RuntimeError(f"数据源 {sheet_name}!{pos} 未找到图片")

    @staticmethod
    def _jmp_format_ref_row(ws, anchor_row, anchor_col, block_height):
        """JMP 格式参考行：从锚点向上跳过最近的一个数据块（block_height 行），
        返回更早一块的最后一行。找不到更早块时回退到锚点上方最近的非空行。"""
        ref = anchor_row - block_height - 1
        if ref >= 1 and ws.cell(row=ref, column=anchor_col).value is not None:
            return ref
        r = anchor_row - 1
        while r >= 1 and ws.cell(row=r, column=anchor_col).value is None:
            r -= 1
        return r if r >= 1 else anchor_row

    def apply_archive_shift_right(self, ws, mapping, data_ws, data_src_ws=None):
        min_row, min_col, max_row, max_col = mapping['block_range']
        header_rows = mapping.get('header_rows', 1)
        new_headers = mapping.get('new_headers', [])
        merged_to_shift = []
        for mr in ws.merged_cells.ranges:
            m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(str(mr))
            if (m_min_row >= min_row and m_max_row <= max_row and
                m_min_col >= min_col and m_max_col <= max_col):
                merged_to_shift.append(str(mr))
        for mr in merged_to_shift:
            ws.unmerge_cells(mr)
        # 右移（纯数值）
        for col in range(max_col, min_col - 1, -1):
            for row in range(min_row, max_row + 1):
                src_cell = data_ws.cell(row=row, column=col)
                dst_cell = ws.cell(row=row, column=col + 1)
                self._warn_error_value(src_cell.value,
                    f"{data_ws.title}!{get_column_letter(col)}{row}",
                    f"{ws.title}!{get_column_letter(col + 1)}{row}")
                dst_cell.value = src_cell.value
                self.copy_cell_style(ws.cell(row=row, column=col), dst_cell)
        for row in range(min_row, max_row + 1):
            ws.cell(row=row, column=min_col).value = None
        for mr in merged_to_shift:
            m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(mr)
            ws.merge_cells(f"{get_column_letter(m_min_col+1)}{m_min_row}:{get_column_letter(m_max_col+1)}{m_max_row}")
        # 扩展覆盖区块首行（含首行上方一行的组表头）的合并单元格，向右扩一格
        self._extend_block_top_merges(ws, min_row, min_col, max_row, max_col)
        # 填充新表头
        for i in range(header_rows):
            row_idx = min_row + i
            cell = ws.cell(row=row_idx, column=min_col)
            if i < len(new_headers):
                cell.value = new_headers[i]
            self.copy_cell_style(ws.cell(row=row_idx, column=min_col + 1), cell)
        # 填充新数据：默认取同一Sheet"来源列"的数值版（data_ws）写入新列
        data_start_row = min_row + header_rows
        src_col = mapping.get('source_col')
        if src_col:
            try:
                src_col_idx = column_index_from_string(str(src_col).upper())
            except ValueError:
                src_col_idx = None
            if src_col_idx:
                for target_row in range(data_start_row, max_row + 1):
                    src_cell = data_ws.cell(row=target_row, column=src_col_idx)
                    dst_cell = ws.cell(row=target_row, column=min_col)
                    self._warn_error_value(src_cell.value,
                        f"{data_ws.title}!{get_column_letter(src_col_idx)}{target_row}",
                        f"{ws.title}!{get_column_letter(min_col)}{target_row}")
                    dst_cell.value = src_cell.value
                    self.copy_cell_style(ws.cell(row=target_row, column=min_col + 1),
                                         ws.cell(row=target_row, column=min_col))
                # 来源列整体为空：通常是文件未由 Excel 保存过，公式缓存值缺失
                data_rows = max_row - data_start_row + 1
                empty_count = sum(
                    1 for r in range(data_start_row, max_row + 1)
                    if data_ws.cell(row=r, column=src_col_idx).value is None
                )
                if empty_count == data_rows:
                    self._fill_warnings.append(
                        f"{ws.title}!{get_column_letter(min_col)}{min_row}:"
                        f"{get_column_letter(min_col)}{max_row} 的来源列 {src_col} 全部为空，"
                        "可能文件未经 Excel 保存（公式缓存值缺失），请先用 Excel 打开保存后再归档。")
                elif empty_count:
                    self._fill_warnings.append(
                        f"{ws.title}!{get_column_letter(min_col)}{min_row}:"
                        f"{get_column_letter(min_col)}{max_row} 的来源列 {src_col} 有 {empty_count}/{data_rows} 行为空。")
        else:
            # 旧配置兼容：从映射的 source_sheet/source_range 读取
            src_ws = data_src_ws if data_src_ws else self.source_wb[mapping['source_sheet']]
            s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(mapping['source_range'])
            for i in range(max_row - data_start_row + 1):
                src_row = s_min_row + i
                if src_row > s_max_row:
                    break
                target_row = data_start_row + i
                src_cell = src_ws.cell(row=src_row, column=s_min_col)
                dst_cell = ws.cell(row=target_row, column=min_col)
                self._warn_error_value(src_cell.value,
                    f"{src_ws.title}!{get_column_letter(s_min_col)}{src_row}",
                    f"{ws.title}!{get_column_letter(min_col)}{target_row}")
                dst_cell.value = src_cell.value
                self.copy_cell_style(ws.cell(row=target_row, column=min_col + 1),
                                     ws.cell(row=target_row, column=min_col))
        # 归档完成后：对数据块（含新列）执行"全部框线 + 粗外框"
        self._apply_block_borders(ws, (min_row, min_col, max_row, max_col + 1))
    def apply_jmp_mapping(self, ws, mapping, src_data_ws=None):
        anchor = mapping['anchor_cell']
        col_letter = ''.join(filter(str.isalpha, anchor))
        row_num = int(''.join(filter(str.isdigit, anchor)))
        start_row = row_num
        start_col = column_index_from_string(col_letter)
        # 表头列数与内容：数据一律从 锚点列 + 表头列数 之后开始张贴
        raw_headers = mapping.get('header_cols', [])
        if isinstance(raw_headers, str):  # 兼容旧配置以字符串存储的情况
            raw_headers = [raw_headers]
        headers = [str(h) for h in raw_headers]
        header_count = len(headers)
        data_start_col = start_col + header_count
        src_sheet_name = mapping['source_sheet']
        src_range = mapping['source_range']
        merge = mapping['merge_columns']
        # 源数据必须取自映射指定的源Sheet的数值版（data_only），不能用目标Sheet代替
        src_ws = src_data_ws if src_data_ws is not None else self.template_wb[src_sheet_name]
        s_min_col, s_min_row, s_max_col, s_max_row = range_boundaries(src_range)
        data_rows = s_max_row - s_min_row + 1
        data_cols = s_max_col - s_min_col + 1
        # 参考行样式：从锚点向上跳过最近的一个数据块（块高=本次写入行数），
        # 取更早一块的格式；找不到时回退到锚点上一行
        block_height = data_rows * data_cols if (merge and data_cols > 1) else data_rows
        ref_row = self._jmp_format_ref_row(ws, start_row, start_col, block_height)
        if merge and data_cols > 1:
            values = []
            for r in range(s_min_row, s_max_row + 1):
                for c in range(s_min_col, s_max_col + 1):
                    val = src_ws.cell(row=r, column=c).value
                    values.append(val)
            row_count = len(values)
            for row_offset in range(row_count):
                current_row = start_row + row_offset
                # 写入表头
                for idx, header_text in enumerate(headers):
                    cell = ws.cell(row=current_row, column=start_col + idx)
                    cell.value = header_text
                # 写入数据
                val = values[row_offset]
                ws.cell(row=current_row, column=data_start_col, value=val)
                # 复制格式（整行）
                for col_idx in range(start_col, data_start_col + 1):
                    ref_cell = ws.cell(row=ref_row, column=col_idx)
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    self.copy_cell_style(ref_cell, target_cell)
        else:
            row_count = data_rows
            target_data_cols = data_cols
            for row_offset in range(row_count):
                current_row = start_row + row_offset
                # 写入表头
                for idx, header_text in enumerate(headers):
                    cell = ws.cell(row=current_row, column=start_col + idx)
                    cell.value = header_text
                # 写入数据
                for c in range(target_data_cols):
                    src_val = src_ws.cell(row=s_min_row + row_offset, column=s_min_col + c).value
                    ws.cell(row=current_row, column=data_start_col + c, value=src_val)
                # 复制格式
                for col_idx in range(start_col, data_start_col + target_data_cols):
                    ref_cell = ws.cell(row=ref_row, column=col_idx)
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    self.copy_cell_style(ref_cell, target_cell)
    @staticmethod
    def copy_cell_style(src, dst):
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)
