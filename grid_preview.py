# -*- coding: utf-8 -*-
"""
方格纸预览生成器 v3
用法: python grid_preview.py [目录路径]
"""

import sys, os, re, webbrowser
from docx import Document

# ── 配置 ──
CELL_SIZE = 32       # px
FONT_SIZE = 20       # px
ROW_GAP = 12         # px
PAGE_PAD_TOP = 18    # mm
PAGE_PAD_BOTTOM = 18 # mm
PAGE_PAD_LEFT = 16   # mm
PAGE_PAD_RIGHT = 16  # mm
COLS = 20
ROWS = 20
CELLS_PER_PAGE = COLS * ROWS


# ═══════════════════════════════════════════
# Step 1: Tokenization
# ═══════════════════════════════════════════

PUNCTS = '\u3002\uff0c\uff1a\uff1b\uff01\uff1f\u3001'  # 。，：；！？、
QUOTES = '\u201c\u201d'


def is_ascii_letter_or_digit(ch):
    """判断是否为 ASCII 字母或数字"""
    o = ord(ch)
    return (48 <= o <= 57) or (65 <= o <= 90) or (97 <= o <= 122)


def tokenize(text):
    """
    将一段文字转为格子列表。
    规则：
      - 连续 ASCII 字母/数字两两配对: "2025"→["20","25"], "Deepseek"→["De","ep","se","ek"]
      - 引号与相邻标点合并一格: 。"→[。"]  "：→["：]  ，"→[，"]
      - 其余每个字符单独一格
    返回: [(content, is_merged), ...]
    """
    tokens = []
    i = 0
    n = len(text)

    while i < n:
        ch = text[i]

        # —— 连续 ASCII 字母/数字：两两配对 ——
        if is_ascii_letter_or_digit(ch):
            run = []
            while i < n and is_ascii_letter_or_digit(text[i]):
                run.append(text[i])
                i += 1
            j = 0
            while j < len(run):
                if j + 1 < len(run):
                    tokens.append((run[j] + run[j+1], False))
                    j += 2
                else:
                    tokens.append((run[j], False))
                    j += 1
            continue

        # —— 引号与相邻标点合并（双字符格） ——
        # 引号在前，标点在后
        if ch in QUOTES and i + 1 < n and text[i+1] in PUNCTS:
            tokens.append((ch + text[i+1], True))
            i += 2
            continue
        # 标点在前，引号在后
        if ch in PUNCTS and i + 1 < n and text[i+1] in QUOTES:
            tokens.append((ch + text[i+1], True))
            i += 2
            continue
        # 引号单独
        if ch in QUOTES:
            tokens.append((ch, False))
            i += 1
            continue

        # —— 普通字符 ——
        tokens.append((ch, False))
        i += 1

    return tokens


# ═══════════════════════════════════════════
# Step 2: 段落分类
# ═══════════════════════════════════════════

def extract_and_classify(filepath):
    doc = Document(filepath)
    paras = []
    first = True
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        kind = 'body'
        if first and ('思想汇报' in text or len(text) <= 6):
            kind = 'title'
            first = False
        elif text.startswith('敬爱的党组织') or text.startswith('亲爱的党组织'):
            kind = 'salutation'
        elif text.startswith('汇报人') or text.startswith('彙報人'):
            kind = 'sign_name'
        elif re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$', text):
            kind = 'sign_date'
        elif text.startswith('以上是') and ('批评指正' in text or '批评、帮助' in text):
            kind = 'closing'

        paras.append({'text': text, 'kind': kind})

    return paras


# ═══════════════════════════════════════════
# Step 3: 布局引擎
# ═══════════════════════════════════════════

def layout_to_pages(paras):
    all_cells = []   # list of rows; each row = list of (content, is_merged)
    current_row = []

    def new_row():
        nonlocal current_row
        while len(current_row) < COLS:
            current_row.append((None, False))
        all_cells.append(current_row)
        current_row = []

    def place_tokens(tokens, align='left', indent=0):
        nonlocal current_row

        if align == 'center':
            total = len(tokens)
            pad = (COLS - total) // 2
            if current_row and any(c[0] is not None for c in current_row):
                new_row()
            for _ in range(pad):
                current_row.append((None, False))
            for t in tokens:
                current_row.append(t)
            while len(current_row) < COLS:
                current_row.append((None, False))
            new_row()

        elif align == 'right':
            if current_row and any(c[0] is not None for c in current_row):
                new_row()
            needed = COLS - len(tokens)
            for _ in range(needed):
                current_row.append((None, False))
            for t in tokens:
                current_row.append(t)
            new_row()

        else:  # left
            if indent > 0:
                if current_row and any(c[0] is not None for c in current_row):
                    new_row()
                for _ in range(indent):
                    current_row.append((None, False))

            for t in tokens:
                if len(current_row) >= COLS:
                    new_row()
                current_row.append(t)

    for p in paras:
        tokens = tokenize(p['text'])
        kind = p['kind']

        if kind == 'title':
            place_tokens(tokens, align='center')
            # 标题后不再额外空行，row-gap 提供间距

        elif kind == 'salutation':
            if current_row and any(c[0] is not None for c in current_row):
                new_row()
            place_tokens(tokens, align='left', indent=0)

        elif kind == 'body':
            place_tokens(tokens, align='left', indent=2)

        elif kind == 'closing':
            place_tokens(tokens, align='left', indent=2)

        elif kind == 'sign_name':
            place_tokens(tokens, align='right')

        elif kind == 'sign_date':
            place_tokens(tokens, align='right')

    if current_row and any(c[0] is not None for c in current_row):
        new_row()

    flat = []
    for row in all_cells:
        flat.extend(row)

    pages = []
    total = len(flat)
    for start in range(0, total, CELLS_PER_PAGE):
        page = flat[start:start + CELLS_PER_PAGE]
        while len(page) < CELLS_PER_PAGE:
            page.append((None, False))
        pages.append(page)

    return pages


# ═══════════════════════════════════════════
# Step 4: HTML 生成
# ═══════════════════════════════════════════

def is_left_punct(ch):
    """判断单字符是否为左偏标点：，、。"""
    return ch in '\u3002\uff0c\u3001'  # 。，、


def is_quote(ch):
    return ch in QUOTES


def generate_html(pages, title_text):
    css = f"""
    :root {{
        --cs: {CELL_SIZE}px;
        --fs: {FONT_SIZE}px;
        --rg: {ROW_GAP}px;
        --pt: {PAGE_PAD_TOP}mm;
        --pb: {PAGE_PAD_BOTTOM}mm;
        --pl: {PAGE_PAD_LEFT}mm;
        --pr: {PAGE_PAD_RIGHT}mm;
    }}
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{
        background: #e8e0d5;
        font-family: "楷体","KaiTi","STKaiti","华文楷体",serif;
        display: flex; flex-wrap: wrap; gap: 30px;
        justify-content: center; padding: 20px;
    }}
    .page {{
        width: 210mm; min-height: 277mm;
        background: #fffef7;
        box-shadow: 0 2px 12px rgba(0,0,0,.3);
        padding: var(--pt) var(--pr) var(--pb) var(--pl);
        position: relative;
    }}
    .grid {{
        display: grid;
        grid-template-columns: repeat({COLS}, var(--cs));
        grid-auto-rows: var(--cs);
        row-gap: var(--rg);
        width: fit-content; margin: 0 auto;
    }}
    .cell {{
        width: var(--cs); height: var(--cs);
        border: .5px solid #d4a8a0;
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs); color: #2a2a2a;
        position: relative; background: #fffef9;
        line-height: 1; overflow: hidden;
    }}
    .cell::after {{
        content: '';
        position: absolute; top: 50%; left: 0; right: 0;
        border-top: .3px dashed #e8ccc4; pointer-events: none;
    }}
    .cell::before {{
        content: '';
        position: absolute; left: 50%; top: 0; bottom: 0;
        border-left: .3px dashed #e8ccc4; pointer-events: none;
    }}
    .cell.empty {{ background: #fffef9; }}

    /* 双字符合并格：两个字符绝对定位重叠 */
    .cell.merged {{
        position: relative;
    }}
    .cell.merged span {{
        position: absolute; inset: 0;
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs); line-height: 1;
    }}
    /* 前引号 \u201c：右上角 */
    .cell.merged span.q-open {{
        justify-content: flex-end; align-items: flex-start;
        padding: 0 1px 0 0;
    }}
    /* 后引号 \u201d：左下角 */
    .cell.merged span.q-close {{
        justify-content: flex-start; align-items: flex-end;
        padding: 0 0 1px 1px;
    }}
    /* 逗号、句号、顿号：左下角 */
    .cell.merged span.p-left {{
        justify-content: flex-start; align-items: flex-end;
        padding: 0 0 1px 1px;
    }}
    /* 冒号、分号、感叹号、问号：居中 */
    .cell.merged span.p-center {{
        justify-content: center; align-items: center;
    }}

    .page-label {{
        text-align: right; font-size: 11px; color: #b88;
        margin-top: 16px; margin-right: 16px;
        font-family: "宋体","SimSun",sans-serif;
    }}
    @media print {{
        body {{ background: white; gap: 0; padding: 0; }}
        .page {{ box-shadow: none; page-break-after: always; margin: 0; }}
        .page:last-child {{ page-break-after: auto; }}
    }}
    """

    total_pages = len(pages)
    pages_html = []
    for pi, page_cells in enumerate(pages):
        cells = []
        for content, is_merged in page_cells:
            if content is None:
                cells.append('<div class="cell empty"></div>')
            elif is_merged:
                ch1 = content[0].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                ch2 = content[1].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                c1 = 'q-open' if content[0] == '\u201c' else ('q-close' if content[0] == '\u201d' else ('p-left' if content[0] in '\u3002\uff0c\u3001' else 'p-center'))
                c2 = 'q-open' if content[1] == '\u201c' else ('q-close' if content[1] == '\u201d' else ('p-left' if content[1] in '\u3002\uff0c\u3001' else 'p-center'))
                cells.append(f'<div class="cell merged"><span class="{c1}">{ch1}</span><span class="{c2}">{ch2}</span></div>')
            else:
                escaped = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                cells.append(f'<div class="cell">{escaped}</div>')

        grid = '<div class="grid">' + ''.join(cells) + '</div>'
        label = f'<div class="page-label">第 {pi+1} 页 | {COLS}×{ROWS} = {CELLS_PER_PAGE} 格 | 共 {total_pages} 页</div>'
        pages_html.append(f'<div class="page">{grid}{label}</div>')

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>{title_text} - 方格纸预览</title>
<style>{css}</style>
</head>
<body>
{''.join(pages_html)}
</body>
</html>"""


# ═══════════════════════════════════════════
# Step 5: 主流程
# ═══════════════════════════════════════════

def process_docx(filepath, outdir=None):
    fname = os.path.basename(filepath)
    name_no_ext = os.path.splitext(fname)[0]
    print(f'  处理: {fname}')

    paras = extract_and_classify(filepath)
    pages = layout_to_pages(paras)

    filled = sum(1 for page in pages for c, _ in page if c is not None)
    print(f'    → {len(paras)} 个段落, {filled} 个有效格, {len(pages)} 页')

    title_text = name_no_ext.replace('【思想汇报】', '').replace('思想汇报', '思想汇报')
    html = generate_html(pages, title_text)

    if outdir is None:
        outdir = os.path.dirname(filepath)
    outpath = os.path.join(outdir, f'{name_no_ext}_方格纸预览.html')
    with open(outpath, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'    → {outpath}')
    return outpath


def main():
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    out_dir = sys.argv[2] if len(sys.argv) > 2 else None

    print(f'方格纸预览生成器 v3')
    print(f'扫描目录: {target_dir}')
    if out_dir:
        print(f'输出目录: {out_dir}')
    print(f'网格: {COLS}×{ROWS}, 格{CELL_SIZE}px, 字{FONT_SIZE}px, 行距{ROW_GAP}px')
    print()

    docx_files = sorted([
        os.path.join(target_dir, f) for f in os.listdir(target_dir)
        if f.endswith('.docx') and not f.startswith('~')
    ])

    if not docx_files:
        print('未找到 .docx 文件。')
        return

    print(f'找到 {len(docx_files)} 个文件\n')
    html_files = []
    for fp in docx_files:
        try:
            out = process_docx(fp, out_dir or target_dir)
            html_files.append(out)
        except Exception as e:
            print(f'    ✗ 错误: {e}')
            import traceback
            traceback.print_exc()

    print(f'\n完成! 共 {len(html_files)} 个预览。')
    if html_files:
        webbrowser.open('file:///' + html_files[0].replace('\\', '/'))


if __name__ == '__main__':
    main()
