# -*- coding: utf-8 -*-
"""
方格纸预览生成器 v3.1 — 支持超格标点
用法: python grid_preview.py [输入目录] [输出目录]
"""

import sys, os, re, webbrowser
from docx import Document

CELL_SIZE = 32
FONT_SIZE = 20
ROW_GAP = 12
PAGE_PAD_TOP = 18
PAGE_PAD_BOTTOM = 18
PAGE_PAD_LEFT = 16
PAGE_PAD_RIGHT = 16
COLS = 20
ROWS = 20
CELLS_PER_PAGE = COLS * ROWS

# 行首禁则
PROHIBITED_START = set('\u3002\uff0c\u3001\uff1a\uff1b\uff01\uff1f\uff09\u3011\u300b\u300f\u300d\u3009\u201d\u2019')
ALLOWED_START = set('\uff08\u007b\u005b\u3010\u300a\u003c\u201c\u2018')


def is_ascii_letter_or_digit(ch):
    o = ord(ch)
    return (48 <= o <= 57) or (65 <= o <= 90) or (97 <= o <= 122)


def tokenize(text):
    PUNCTS = '\u3002\uff0c\uff1a\uff1b\uff01\uff1f\u3001'
    QUOTES = '\u201c\u201d'
    tokens = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if is_ascii_letter_or_digit(ch):
            run = []
            while i < n and is_ascii_letter_or_digit(text[i]):
                run.append(text[i]); i += 1
            j = 0
            while j < len(run):
                if j + 1 < len(run):
                    tokens.append((run[j] + run[j+1], False))
                    j += 2
                else:
                    tokens.append((run[j], False))
                    j += 1
            continue
        if ch in QUOTES and i + 1 < n and text[i+1] in PUNCTS:
            tokens.append((ch + text[i+1], True)); i += 2; continue
        if ch in PUNCTS and i + 1 < n and text[i+1] in QUOTES:
            tokens.append((ch + text[i+1], True)); i += 2; continue
        if ch in QUOTES:
            tokens.append((ch, False)); i += 1; continue
        tokens.append((ch, False)); i += 1
    return tokens


def extract_and_classify(filepath):
    doc = Document(filepath)
    paras = []
    first = True
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text: continue
        kind = 'body'
        if first and ('\u601d\u60f3\u6c47\u62a5' in text or len(text) <= 6):
            kind = 'title'; first = False
        elif text.startswith('\u656c\u7231\u7684\u515a\u7ec4\u7ec7') or text.startswith('\u4eb2\u7231\u7684\u515a\u7ec4\u7ec7'):
            kind = 'salutation'
        elif text.startswith('\u6c47\u62a5\u4eba') or text.startswith('\u5f59\u5831\u4eba'):
            kind = 'sign_name'
        elif re.match(r'^\d{4}\u5e74\d{1,2}\u6708\d{1,2}\u65e5$', text):
            kind = 'sign_date'
        elif text.startswith('\u4ee5\u4e0a\u662f') and ('\u6279\u8bc4\u6307\u6b63' in text or '\u6279\u8bc4\u3001\u5e2e\u52a9' in text):
            kind = 'closing'
        paras.append({'text': text, 'kind': kind})
    return paras


def should_overflow(content):
    ch = content[0]
    if ch in ALLOWED_START: return False
    if ch == '\u2014': return False
    if ch in PROHIBITED_START: return True
    return False


def has_opening_quote(content):
    return len(content) == 2 and (content[0] in '\u201c\u2018' or content[1] in '\u201c\u2018')


def layout_to_pages(paras):
    all_rows = []
    current_row = []

    def flush_row(overflow=None):
        nonlocal current_row
        while len(current_row) < COLS:
            current_row.append((None, False))
        all_rows.append({'cells': current_row, 'overflow': overflow})
        current_row = []

    def place_tokens(tokens, align='left', indent=0):
        nonlocal current_row

        if align == 'center':
            total = len(tokens)
            pad = (COLS - total) // 2
            if current_row and any(c[0] is not None for c in current_row):
                flush_row()
            for _ in range(pad):
                current_row.append((None, False))
            for t in tokens:
                current_row.append(t)
            flush_row()

        elif align == 'right':
            if current_row and any(c[0] is not None for c in current_row):
                flush_row()
            needed = COLS - len(tokens)
            for _ in range(needed):
                current_row.append((None, False))
            for t in tokens:
                current_row.append(t)
            flush_row()

        else:
            if indent > 0:
                if current_row and any(c[0] is not None for c in current_row):
                    flush_row()
                for _ in range(indent):
                    current_row.append((None, False))

            for t in tokens:
                if len(current_row) >= COLS:
                    content, is_merged = t
                    if should_overflow(content):
                        if is_merged and has_opening_quote(content):
                            if content[0] in '\u201c\u2018':
                                ov_ch, nl_ch = content[1], content[0]
                            else:
                                ov_ch, nl_ch = content[0], content[1]
                            flush_row(overflow=(ov_ch, False))
                            current_row.append((nl_ch, False))
                        else:
                            flush_row(overflow=t)
                        continue
                    else:
                        flush_row()
                current_row.append(t)

    for p in paras:
        tokens = tokenize(p['text'])
        kind = p['kind']
        if kind == 'title':
            place_tokens(tokens, align='center')
        elif kind == 'salutation':
            if current_row and any(c[0] is not None for c in current_row):
                flush_row()
            place_tokens(tokens, align='left', indent=0)
        elif kind in ('body', 'closing'):
            place_tokens(tokens, align='left', indent=2)
        elif kind in ('sign_name', 'sign_date'):
            place_tokens(tokens, align='right')

    if current_row and any(c[0] is not None for c in current_row):
        flush_row()

    # 分页
    slots_per_page = (COLS + 1) * ROWS
    flat = []
    for row in all_rows:
        for c in row['cells']: flat.append(c)
        flat.append(row['overflow'])

    pages = []
    for start in range(0, len(flat), slots_per_page):
        chunk = flat[start:start + slots_per_page]
        page_rows = []
        for r in range(ROWS):
            base = r * (COLS + 1)
            if base >= len(chunk): break
            cells = chunk[base:base + COLS]
            ov = chunk[base + COLS] if base + COLS < len(chunk) else None
            while len(cells) < COLS:
                cells.append((None, False))
            page_rows.append({'cells': cells, 'overflow': ov})
        pages.append(page_rows)

    return pages


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
        font-family: "\u6977\u4f53","KaiTi","STKaiti","\u534e\u6587\u6977\u4f53",serif;
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
    .rows {{
        display: flex; flex-direction: column; gap: var(--rg);
        width: fit-content; margin: 0 auto;
    }}
    .row {{
        display: flex; align-items: flex-start;
    }}
    .grid-20 {{
        display: grid;
        grid-template-columns: repeat({COLS}, var(--cs));
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
    .overflow-cell {{
        width: var(--cs); height: var(--cs);
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs); color: #2a2a2a;
        line-height: 1;
    }}
    .overflow-cell.empty {{ visibility: hidden; }}
    .cell.merged {{
        position: relative;
    }}
    .cell.merged span {{
        position: absolute; inset: 0;
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs); line-height: 1;
    }}
    .cell.merged span.p-left {{ padding: 0px 4px 0px 0px; }}
    .cell.merged span.p-right {{ padding: 0px 0px 0px 18px; }}
    .page-label {{
        text-align: right; font-size: 11px; color: #b88;
        margin-top: 16px; margin-right: 16px;
        font-family: "\u5b8b\u4f53","SimSun",sans-serif;
    }}
    @media print {{
        body {{ background: white; gap: 0; padding: 0; }}
        .page {{ box-shadow: none; page-break-after: always; margin: 0; }}
        .page:last-child {{ page-break-after: auto; }}
    }}
    """

    def cell_html(c_tuple):
        content, is_merged = c_tuple
        if content is None:
            return '<div class="cell empty"></div>'
        if is_merged:
            ch1 = content[0].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
            ch2 = content[1].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
            return f'<div class="cell merged"><span class="p-left">{ch1}</span><span class="p-right">{ch2}</span></div>'
        escaped = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
        return f'<div class="cell">{escaped}</div>'

    def overflow_html(ov_tuple):
        if ov_tuple is None or ov_tuple[0] is None:
            return '<div class="overflow-cell empty"></div>'
        content, is_merged = ov_tuple
        escaped = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
        return f'<div class="overflow-cell">{escaped}</div>'

    total_pages = len(pages)
    pages_html = []
    for pi, page_rows in enumerate(pages):
        rows_html = []
        for row in page_rows:
            cells = ''.join(cell_html(c) for c in row['cells'])
            ov = overflow_html(row['overflow'])
            rows_html.append(f'<div class="row"><div class="grid-20">{cells}</div>{ov}</div>')
        label = f'<div class="page-label">第 {pi+1} 页 | {COLS}\u00d7{ROWS} = {CELLS_PER_PAGE} 格 | 共 {total_pages} 页</div>'
        pages_html.append(f'<div class="page"><div class="rows">{"".join(rows_html)}</div>{label}</div>')

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>{title_text}</title>
<style>{css}</style>
</head>
<body>
{''.join(pages_html)}
</body>
</html>"""


def process_docx(filepath, outdir=None):
    fname = os.path.basename(filepath)
    name_no_ext = os.path.splitext(fname)[0]
    print(f'  处理: {fname}')
    paras = extract_and_classify(filepath)
    pages = layout_to_pages(paras)
    filled = sum(1 for page in pages for row in page for c, _ in row['cells'] if c is not None)
    print(f'    \u2192 {len(paras)} 个段落, {filled} 个有效格, {len(pages)} 页')
    html = generate_html(pages, name_no_ext)
    if outdir is None:
        outdir = os.path.dirname(filepath)
    outpath = os.path.join(outdir, f'{name_no_ext}_\u65b9\u683c\u7eb8\u9884\u89c8.html')
    with open(outpath, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'    \u2192 {outpath}')
    return outpath


def main():
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    out_dir = sys.argv[2] if len(sys.argv) > 2 else None
    print(f'\u65b9\u683c\u7eb8\u9884\u89c8\u751f\u6210\u5668 v3.1')
    print(f'\u626b\u63cf\u76ee\u5f55: {target_dir}')
    if out_dir:
        print(f'\u8f93\u51fa\u76ee\u5f55: {out_dir}')
    print(f'\u7f51\u683c: {COLS}\u00d7{ROWS}, \u683c{CELL_SIZE}px, \u5b57{FONT_SIZE}px, \u884c\u8ddd{ROW_GAP}px\n')
    docx_files = sorted([os.path.join(target_dir, f) for f in os.listdir(target_dir) if f.endswith('.docx') and not f.startswith('~')])
    if not docx_files:
        print('\u672a\u627e\u5230 .docx \u6587\u4ef6\u3002')
        return
    print(f'\u627e\u5230 {len(docx_files)} \u4e2a\u6587\u4ef6\n')
    html_files = []
    for fp in docx_files:
        try:
            out = process_docx(fp, out_dir or target_dir)
            html_files.append(out)
        except Exception as e:
            print(f'    \u2717 \u9519\u8bef: {e}')
            import traceback; traceback.print_exc()
    print(f'\n\u5b8c\u6210! \u5171 {len(html_files)} \u4e2a\u9884\u89c8\u3002')
    if html_files:
        webbrowser.open('file:///' + html_files[0].replace('\\', '/'))


if __name__ == '__main__':
    main()
