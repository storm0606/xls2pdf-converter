"""
XLS → PDF 核心转换模块 v5 — 像素级还原 Excel 原样式

v5 修复：
  - 对齐映射修正（0=general/1=left/2=center/3=right）
  - CJK 字体回退（英文字体用于中文时自动切换）
  - 表格整体居中
  - 字体注册名使用稳定哈希

v5.1 修复：
  - 恢复排除 #000000 背景（xlrd无填充标记 pattern_colour_index=64 映射为 #000000）
"""
import os
import glob
import hashlib
import tempfile

import xlrd
from reportlab.lib.colors import HexColor, black
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from font_subsetter import find_font, subset_font

# Bold font suffix mapping (font_name lower → bold font filename)
_BOLD_FONT_MAP = {
    'arial': 'arialbd.ttf',
    'liberation sans': 'LiberationSans-Bold.ttf',
}

# xlrd BIFF8 默认 64 色调色板（仅作为 fallback）
_DEFAULT_PALETTE = [
    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
    '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
    '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
    '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
    '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
    '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
    '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
]

# 已知的中文字体名称白名单（这些字体包含CJK字形，不需要回退）
_CJK_FONT_NAMES = {
    '微软雅黑', 'microsoft yahei', '宋体', 'simsun', '黑体', 'simhei',
    '仿宋', 'fangsong', '楷体', 'kaiti',
    'wenquanyi micro hei', 'wenquanyi zen hei',
    'noto sans cjk', 'noto serif cjk',
    'microsoft jhenghei', 'microsoft yahei ui',
    'pingfang sc', 'hiragino sans gb', 'stheiti',
    'source han sans', 'source han serif',
}


def _colour_index_to_hex(wb, idx):
    """使用 wb.colour_map 获取真实颜色（支持自定义调色板）"""
    if idx is None or idx < 0:
        return '#000000'
    # 优先用 wb.colour_map（包含自定义调色板）
    if hasattr(wb, 'colour_map') and idx in wb.colour_map:
        rgb = wb.colour_map[idx]
        if rgb is not None:
            return '#{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])
    # Fallback 到默认调色板
    if idx < len(_DEFAULT_PALETTE):
        return _DEFAULT_PALETTE[idx]
    return '#000000'


def _xlrd_date_to_str(val, wb):
    try:
        dt = xlrd.xldate_as_tuple(val, wb.datemode)
        if dt[3] == 0 and dt[4] == 0 and dt[5] == 0:
            return f'{dt[0]}/{dt[1]:02d}/{dt[2]:02d}'
        return f'{dt[0]}/{dt[1]:02d}/{dt[2]:02d} {dt[3]:02d}:{dt[4]:02d}:{dt[5]:02d}'
    except Exception:
        return str(val)


# xlrd 列宽单位：1/256 字符宽度
# 1 字符宽度 ≈ 7 像素 (96 DPI) = 7*72/96 = 5.25 pt
# 1 单位 = 5.25 / 256 ≈ 0.0205 pt
COL_UNIT_PT = 7.0 / 256.0 * 72.0 / 96.0
BORDER_THIN = 0.5   # 细线
BORDER_MEDIUM = 1.0  # 中等线
BORDER_BLACK = 1.0   # 黑色边框专用线宽（加粗黑色分隔线）
MARGIN = 14
A4_W = 595.28
A4_H = 841.89
USABLE_H = A4_H - MARGIN * 2


def parse_xls(xls_path):
    wb = xlrd.open_workbook(xls_path, formatting_info=True)
    sheets = []

    for sidx in range(wb.nsheets):
        ws = wb.sheet_by_index(sidx)

        # 合并单元格
        merged_map = {}
        merged_slave = set()
        for rlo, rhi, clo, chi in ws.merged_cells:
            merged_map[(rlo, clo)] = (rhi, chi)
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if r != rlo or c != clo:
                        merged_slave.add((r, c))

        # 列宽
        col_widths = []
        for ci in range(ws.ncols):
            cw = ws.colinfo_map.get(ci)
            col_widths.append(cw.width if cw else 2962)

        # 行高
        row_heights = {}
        for ri in range(ws.nrows):
            rh = ws.rowinfo_map.get(ri)
            if rh is not None and rh.height > 0:
                row_heights[ri] = rh.height

        # 单元格样式
        cell_styles = {}
        for ri in range(ws.nrows):
            for ci in range(ws.ncols):
                cell = ws.cell(ri, ci)
                xf = wb.xf_list[cell.xf_index]
                font = wb.font_list[xf.font_index]

                text = ''
                if cell.ctype == 1:
                    text = str(cell.value)
                elif cell.ctype == 2:
                    v = cell.value
                    text = str(int(v)) if v == int(v) else str(v)
                elif cell.ctype == 3:
                    text = _xlrd_date_to_str(cell.value, wb)
                elif cell.ctype == 4:
                    text = 'TRUE' if cell.value else 'FALSE'

                cell_styles[(ri, ci)] = {
                    'text': text,
                    'ctype': cell.ctype,
                    'font_name': font.name,
                    'font_height_pt': font.height / 20.0,
                    'font_bold': font.weight >= 700,
                    'font_colour_hex': _colour_index_to_hex(wb, font.colour_index),
                    'horz_align': xf.alignment.hor_align,
                    'vert_align': xf.alignment.vert_align,
                    'border_left': xf.border.left_line_style,
                    'border_right': xf.border.right_line_style,
                    'border_top': xf.border.top_line_style,
                    'border_bottom': xf.border.bottom_line_style,
                    'border_left_colour': _colour_index_to_hex(wb, xf.border.left_colour_index),
                    'border_right_colour': _colour_index_to_hex(wb, xf.border.right_colour_index),
                    'border_top_colour': _colour_index_to_hex(wb, xf.border.top_colour_index),
                    'border_bottom_colour': _colour_index_to_hex(wb, xf.border.bottom_colour_index),
                    'bg_colour_hex': _colour_index_to_hex(wb, xf.background.pattern_colour_index),
                }

        sheets.append({
            'name': ws.name,
            'nrows': ws.nrows,
            'ncols': ws.ncols,
            'col_widths': col_widths,
            'row_heights': row_heights,
            'merged_map': merged_map,
            'merged_slave': merged_slave,
            'cell_styles': cell_styles,
        })

    return sheets, wb


def _collect_chars(sheets):
    chars = set()
    for sheet in sheets:
        for style in sheet['cell_styles'].values():
            if style['text']:
                chars.update(style['text'])
    chars.update(' !"#$%&\'()*+,-./0123456789:;<=>?@')
    chars.update('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    chars.update('abcdefghijklmnopqrstuvwxyz')
    chars.update('[\\]^_`{|}~')
    return chars


def _has_cjk(text):
    """检查文本是否包含中日韩(CJK)字符，覆盖常用范围"""
    if not text:
        return False
    for ch in text:
        if ('\u4e00' <= ch <= '\u9fff'      # CJK Unified Ideographs
            or '\u3000' <= ch <= '\u303f'    # CJK Symbols and Punctuation
            or '\uff00' <= ch <= '\uffef'    # Fullwidth Forms
            or '\u3040' <= ch <= '\u309f'    # Hiragana
            or '\u30a0' <= ch <= '\u30ff'    # Katakana
            or '\uac00' <= ch <= '\ud7af'    # Hangul Syllables
            or '\u3400' <= ch <= '\u4dbf'):  # CJK Extension A
            return True
    return False


def _is_cjk_font(font_name):
    """判断字体名是否属于已知CJK字体（含CJK字形，无需回退）"""
    return font_name.lower().strip() in _CJK_FONT_NAMES


def _stable_font_id(font_name):
    """生成稳定的字体注册名后缀（替代 id() 避免地址复用碰撞）"""
    return hashlib.md5(font_name.encode('utf-8')).hexdigest()[:8]


def _register_subsets(sheets):
    """Register regular and bold font subsets. Returns (font_map, bold_font_map, fallback_reg, _)."""
    chars = _collect_chars(sheets)
    font_names = set()
    for sheet in sheets:
        for style in sheet['cell_styles'].values():
            if style['font_name']:
                font_names.add(style['font_name'])

    registered = {}
    bold_registered = {}
    fallback_reg = None
    bold_fallback_reg = None

    for fname in font_names:
        fpath = find_font(fname)
        if not fpath:
            print(f'[WARN] Font not found: {fname}')
            continue
        try:
            subset_path = subset_font(fpath, chars, font_index=0)
        except Exception as e:
            print(f'[WARN] Subset failed for {fname}: {e}')
            subset_path = fpath
        reg_name = f'f_{_stable_font_id(fname)}'
        try:
            pdfmetrics.registerFont(TTFont(reg_name, subset_path))
            registered[fname] = reg_name
        except Exception as e:
            print(f'[WARN] Register font failed for {fname}: {e}')

        # ── 注册粗体字体变体 ──
        bold_fname = _BOLD_FONT_MAP.get(fname.lower())
        if bold_fname:
            bold_fpath = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 'fonts', bold_fname
            )
            if os.path.exists(bold_fpath):
                try:
                    bold_subset_path = subset_font(bold_fpath, chars, font_index=0)
                    bold_reg_name = f'f_{_stable_font_id(fname)}_bold'
                    pdfmetrics.registerFont(TTFont(bold_reg_name, bold_subset_path))
                    bold_registered[fname] = bold_reg_name
                    print(f'[INFO] Bold font registered: {bold_fname}')
                except Exception as e:
                    print(f'[WARN] Register bold font failed for {bold_fname}: {e}')

    # ── 额外注册中文字体作为 fallback（用于英文字体+中文文本的情况） ──
    fallback_name = '微软雅黑'
    fpath = find_font(fallback_name)
    if fpath:
        try:
            fallback_path = subset_font(fpath, chars, font_index=0)
            fallback_reg = 'fallback_cjk_sub'
            pdfmetrics.registerFont(TTFont(fallback_reg, fallback_path))
            print(f'[INFO] CJK fallback font registered: {fallback_name}')
            # 注册粗体 fallback（使用同字体，配合描边模拟）
            bold_fallback_reg = None
        except Exception as e:
            print(f'[WARN] CJK fallback font registration failed: {e}')

    if not fallback_reg:
        for cand_glob in [
            '/usr/share/fonts/**/wqy-microhei*',
            '/usr/share/fonts/**/WenQuanYi*',
            '/usr/share/fonts/**/NotoSansCJK*',
            '/usr/share/fonts/**/NotoSerifCJK*',
        ]:
            flist = glob.glob(cand_glob, recursive=True)
            if flist:
                try:
                    fallback_path = subset_font(flist[0], chars, font_index=0)
                    fallback_reg = 'fallback_cjk_sub'
                    pdfmetrics.registerFont(TTFont(fallback_reg, fallback_path))
                    print(f'[INFO] CJK fallback font registered: {flist[0]}')
                    bold_fallback_reg = None
                except Exception as e:
                    print(f'[WARN] CJK fallback reg failed ({flist[0]}): {e}')
                break

    return registered, bold_registered, fallback_reg, bold_fallback_reg


def _draw_cell(c, ri, ci, sheet, col_x, row_y, col_w, row_h, font_map, bold_font_map=None, fallback_reg=None, bold_fallback_reg=None):
    """绘制单个单元格（背景+边框+文本）"""
    style = sheet['cell_styles'].get((ri, ci))
    if not style:
        return

    # 确定绘制区域
    if (ri, ci) in sheet['merged_map']:
        rhi, chi = sheet['merged_map'][(ri, ci)]
        x1 = col_x[ci]
        x2 = col_x[chi - 1] + col_w[chi - 1] if chi - 1 < len(col_w) else x1
        y1 = row_y[ri]
        y2 = row_y[rhi - 1] - row_h[rhi - 1] if rhi - 1 < len(row_h) else y1
    else:
        x1 = col_x[ci]
        x2 = col_x[ci] + col_w[ci]
        y1 = row_y[ri]
        y2 = row_y[ri] - row_h[ri]

    w = x2 - x1
    h = y1 - y2
    if w <= 0 or h <= 0:
        return

    # ── 背景 ──
    bg = style['bg_colour_hex']
    # 排除 #FFFFFF（默认白色）和 #000000（xlrd无填充标记 pattern_colour_index=64 的映射值）
    if bg and bg not in ('#FFFFFF', '#000000'):
        try:
            c.setFillColor(HexColor(bg))
            c.rect(x1, y2, w, h, fill=1, stroke=0)
        except Exception:
            pass

    # ── 边框 ──
    for side, bw, bc in [
        ('left',   style['border_left'],   style['border_left_colour']),
        ('right',  style['border_right'],  style['border_right_colour']),
        ('top',    style['border_top'],    style['border_top_colour']),
        ('bottom', style['border_bottom'], style['border_bottom_colour']),
    ]:
        if bw > 0:
            try:
                c.setStrokeColor(HexColor(bc))
            except Exception:
                c.setStrokeColor(black)
            is_black_border = (bc == '#000000')
            if is_black_border:
                c.setLineWidth(BORDER_BLACK)
            else:
                c.setLineWidth(BORDER_THIN if bw == 1 else BORDER_MEDIUM)
            if side == 'left':
                c.line(x1, y2, x1, y1)
            elif side == 'right':
                c.line(x2, y2, x2, y1)
            elif side == 'top':
                c.line(x1, y1, x2, y1)
            elif side == 'bottom':
                c.line(x1, y2, x2, y2)

    # ── 文本 ──
    text = style['text']
    if not text:
        return

    font_name = style['font_name']
    is_bold = style['font_bold']

    # ── CJK 字体回退：非CJK字体遇到CJK文本时切换到中文字体 ──
    needs_cjk_fallback = _has_cjk(text) and not _is_cjk_font(font_name)
    is_cjk = _is_cjk_font(font_name)

    if needs_cjk_fallback:
        base_reg = fallback_reg or font_map.get(font_name, 'Helvetica')
        bold_reg = bold_fallback_reg if is_bold else None
    elif is_cjk:
        # CJK 字体：常规字体直接使用，粗体使用 fallback 粗体（同字体，stroke模拟）
        base_reg = font_map.get(font_name, 'Helvetica')
        bold_reg = bold_fallback_reg if is_bold and bold_fallback_reg else None
    else:
        base_reg = font_map.get(font_name, 'Helvetica')
        bold_reg = bold_font_map.get(font_name) if is_bold and bold_font_map else None

    # 选择最终字体注册名
    if is_bold and bold_reg:
        # 方案A：有专门的粗体字体 → 直接使用
        reg_name = bold_reg
        use_stroke_sim = False
    elif is_bold:
        # 方案B：无粗体字体 → 使用常规字体 + 描边模拟
        reg_name = base_reg
        use_stroke_sim = True
    else:
        reg_name = base_reg
        use_stroke_sim = False

    font_size = style['font_height_pt']
    if font_size < 4:
        font_size = 10

    try:
        c.setFillColor(HexColor(style['font_colour_hex']))
    except Exception:
        c.setFillColor(black)

    try:
        c.setFont(reg_name, font_size)
    except Exception:
        try:
            c.setFont('Helvetica', font_size)
            reg_name = 'Helvetica'
        except Exception:
            return

    horz = style['horz_align']
    vert = style['vert_align']
    pad = 2

    # 水平对齐：0=general, 1=left, 2=center, 3=right
    if horz == 0:
        ctype = style.get('ctype', -1)
        if ctype == 2:
            tw = c.stringWidth(text, reg_name, font_size)
            tx = x2 - tw - pad
        else:
            tx = x1 + pad
    elif horz == 1:
        tx = x1 + pad
    elif horz == 2:
        tw = c.stringWidth(text, reg_name, font_size)
        tx = x1 + (w - tw) / 2.0
    elif horz == 3:
        tw = c.stringWidth(text, reg_name, font_size)
        tx = x2 - tw - pad
    else:
        tx = x1 + pad

    if vert == 2:
        ty = y2 + pad
    elif vert == 1:
        ty = y2 + (h - font_size) / 2.0
    else:
        ty = y1 - font_size - pad

    # ── 渲染文本（含加粗处理） ──
    if use_stroke_sim:
        # 方案B：描边模拟加粗（CJK粗体无专用字体时的回退）
        c.saveState()
        c.setLineWidth(0.3)
        c.drawString(tx + 0.3, ty, text)
        c.restoreState()
    c.drawString(tx, ty, text)


def xls_to_pdf(xls_path, output_path=None):
    sheets, wb = parse_xls(xls_path)
    font_map, bold_font_map, fallback_reg, bold_fallback_reg = _register_subsets(sheets)

    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix='.pdf')
        os.close(fd)

    # Calculate max page width across all sheets
    max_page_w = A4_W
    for sheet in sheets:
        if sheet['nrows'] == 0:
            continue
        col_w = [w * COL_UNIT_PT for w in sheet['col_widths']]
        total_w = sum(col_w) + MARGIN * 2
        max_page_w = max(max_page_w, total_w)

    c = canvas.Canvas(output_path, pagesize=(max_page_w, A4_H))

    for sheet in sheets:
        ncols = sheet['ncols']
        nrows = sheet['nrows']

        if nrows == 0:
            continue

        col_w = [w * COL_UNIT_PT for w in sheet['col_widths']]
        row_h = [sheet['row_heights'].get(ri, 300) / 20.0 for ri in range(nrows)]

        # ── 计算表格整体居中偏移 ──
        total_table_w = sum(col_w)
        page_w = max_page_w  # 使用与Canvas一致的实际页面宽度
        if total_table_w < page_w - 2 * MARGIN:
            offset_x = (page_w - total_table_w) / 2.0
        else:
            offset_x = MARGIN

        # 分页
        pages = []
        cur_row = 0
        while cur_row < nrows:
            used_h = 0
            page_start = cur_row
            while cur_row < nrows:
                if used_h + row_h[cur_row] > USABLE_H and cur_row > page_start:
                    break
                used_h += row_h[cur_row]
                cur_row += 1
            pages.append((page_start, cur_row))

        for page_start, page_end in pages:
            row_y = {}
            y = A4_H - MARGIN
            for ri in range(page_start, page_end):
                row_y[ri] = y
                y -= row_h[ri]

            col_x = {}
            x = offset_x
            for ci in range(ncols):
                col_x[ci] = x
                x += col_w[ci]

            for ri in range(page_start, page_end):
                for ci in range(ncols):
                    if (ri, ci) in sheet['merged_slave']:
                        continue
                    if (ri, ci) in sheet['merged_map']:
                        rhi, chi = sheet['merged_map'][(ri, ci)]
                        if rhi <= page_start or ri >= page_end:
                            continue
                    _draw_cell(c, ri, ci, sheet, col_x, row_y, col_w, row_h,
                               font_map, bold_font_map=bold_font_map,
                               fallback_reg=fallback_reg, bold_fallback_reg=bold_fallback_reg)

            c.showPage()

    c.save()

    return output_path
