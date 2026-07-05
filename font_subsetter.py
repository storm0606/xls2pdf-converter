"""
字体子集化模块 — 只嵌入文档中实际使用到的字符
"""
import os
import tempfile
from fontTools.subset import Subsetter, Options
from fontTools.ttLib import TTFont

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')

# 字体名称到文件名的映射
_FONT_NAME_MAP = {
    '微软雅黑': 'msyh.ttf',
    'microsoft yahei': 'msyh.ttf',
    'arial': 'arial.ttf',
    'arial unicode ms': 'arialuni.ttf',
    'simsun': 'simsun.ttc',
    '宋体': 'simsun.ttc',
    'simhei': 'simhei.ttc',
    '黑体': 'simhei.ttc',
}

# 字体fallback映射（原始字体找不到时用替代字体）
_FONT_FALLBACK = {
    '微软雅黑': 'WenQuanYi Micro Hei',
    'microsoft yahei': 'WenQuanYi Micro Hei',
    'arial': 'WenQuanYi Micro Hei',
    'simsun': 'WenQuanYi Micro Hei',
    '宋体': 'WenQuanYi Micro Hei',
    'simhei': 'WenQuanYi Micro Hei',
    '黑体': 'WenQuanYi Micro Hei',
}

# 已知fallback字体路径
_FALLBACK_FONT_PATHS = [
    '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
    '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
]

# 系统字体搜索目录
_SYS_FONT_DIRS = [
    '/usr/share/fonts',
    '/usr/local/share/fonts',
    '/usr/share/fonts/truetype',
    '/usr/share/fonts/opentype',
]


def find_font(font_name):
    """查找字体文件，返回绝对路径。找不到时自动fallback到系统中文字体。"""
    name_lower = font_name.lower().strip()
    fname = _FONT_NAME_MAP.get(name_lower, f'{font_name}.ttf')

    # 1) 本地 fonts/ 目录
    local = os.path.join(FONT_DIR, fname)
    if os.path.exists(local):
        return local

    # 2) 系统字体目录 — 精确匹配
    for d in _SYS_FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if f.lower() == fname.lower():
                    return os.path.join(root, f)

    # 3) 系统字体目录 — 模糊匹配
    for d in _SYS_FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if name_lower.replace(' ', '') in f.lower().replace(' ', ''):
                    return os.path.join(root, f)

    # 4) Fallback到替代字体
    fallback_name = _FONT_FALLBACK.get(name_lower)
    if fallback_name:
        fallback_lower = fallback_name.lower()
        for d in _SYS_FONT_DIRS:
            if not os.path.isdir(d):
                continue
            for root, _, files in os.walk(d):
                for f in files:
                    if fallback_lower.replace(' ', '') in f.lower().replace(' ', ''):
                        return os.path.join(root, f)

    # 5) 最后尝试已知fallback路径
    for p in _FALLBACK_FONT_PATHS:
        if os.path.exists(p):
            return p

    return None


def subset_font(font_path, chars, output_path=None, font_index=0):
    """
    对字体做子集化，只保留用到的字符。
    返回子集化后的字体文件路径。
    """
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix='.ttf')
        os.close(fd)

    options = Options()
    options.layout_features = ['*']
    options.name_IDs = ['*']
    options.legacy_cmap = True
    options.symbol_cmap = True
    options.glyph_names = True

    # TTC 需要先提取单个字体
    actual_font_path = font_path
    if font_path.lower().endswith('.ttc'):
        ttfont = TTFont(font_path, fontNumber=font_index)
        tmp_ttf = output_path + '.extracted.ttf'
        ttfont.save(tmp_ttf)
        ttfont.close()
        actual_font_path = tmp_ttf

    subsetter = Subsetter(options=options)
    font = TTFont(actual_font_path)
    subsetter.populate(text=''.join(sorted(set(chars))))
    subsetter.subset(font)
    font.save(output_path)
    font.close()

    # 清理临时提取的TTF
    if actual_font_path != font_path and os.path.exists(actual_font_path):
        os.unlink(actual_font_path)

    return output_path
