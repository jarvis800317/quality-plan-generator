# python-docx 技術規格與已知 Bug 修正

---

## ⭐ 零、品質計畫文件格式強制規定（每次撰寫腳本前必讀）

> 以下為品質計畫的**通用格式規定**，每次產生腳本時必須逐條實作，不得遺漏。

### 0-1 頁尾格式（三段式分節）

| 文件區域 | 頁尾格式 |
|---------|---------|
| 封面 | **無頁尾**（不顯示任何頁碼） |
| 目錄 | 羅馬數字小寫，水平置中：i、ii、iii、iv… |
| 內文（第壹/一章起） | 阿拉伯數字，水平置中：1、2、3、4… |

實作函式：`insert_section_break()`、`setup_section_footer()`（見 §七）

### 0-2 目錄格式

- 呈現**標題 1（章）**與**標題 2（節）**所在頁數
- 頁碼為阿拉伯數字，**靠右對齊**（Tab 定位點 + 點線引導至右邊界）
- 章標題加粗，節標題一般粗細
- 實作函式：`make_toc()`（見 §四 → make_toc）

### 0-3 顏色規定（全文強制黑色）

| 元素 | 規定 |
|-----|------|
| 所有文字（封面/目錄/標題/內文/附件） | **黑色 RGB(0,0,0)**，禁止藍色或彩色 |
| 表格框線 | **黑色** |
| 表格標題列背景 | 淺灰 `#F2F2F2`（禁止藍色/彩色） |
| 表格資料列背景 | 白色 `#FFFFFF`（禁止條紋彩色） |
| 流程圖形狀底色 | **無填色**（白色），外框線黑色 |
| PIL 圖形文字/框線 | 黑色 `(0,0,0)` |

> `HEADING_COLORS` 必須為 `(0,0,0)`；`make_three_col_flow()` 標題/框線/欄頭全黑。

### 0-4 字型彈性規定

- 每次產出前**必須分析**當次監造計畫的字型（中文/英數/各層級大小）
- 腳本中**全部使用常數**：`DOCX_FONT_CJK`、`DOCX_FONT_LATIN`、`DOCX_SZ_*`
- **禁止**在函式內直接寫死任何字型名稱或字型大小數值
- 封面、目錄、標題、內文、附件各層級的字型/大小，需與監造計畫對應層級相近

### 0-5 表格編號規定

- 表格本身水平置中（`WD_TABLE_ALIGNMENT.CENTER`）
- **表名位於表格上方**，水平置中
- 格式：`表{章號}-{序號} 表名`（章號為**阿拉伯數字**，壹=1、貳=2…）
  - 例：`表1-1 工程主要施工項目及數量表`、`表2-1 工作職掌表`
- 實作：先呼叫 `add_table_caption(doc, chapter, seq, name)`，再呼叫 `make_table()`（見 §八）

### 0-6 圖形編號規定

- 圖形段落水平置中（`WD_ALIGN_PARAGRAPH.CENTER`）
- **圖名位於圖形下方**，水平置中
- 格式：`圖{章號}-{序號} 圖名`（章號為**阿拉伯數字**）
  - 例：`圖2-1 品管組織架構圖`、`圖5-1 材料/設備選定送審流程圖`
- 實作：先呼叫 `add_img()`，再呼叫 `add_figure_caption(doc, chapter, seq, name)`（見 §八）

### 0-7 換頁規定

| 情境 | 規定 |
|------|------|
| **標題 1（章）** | **新章開始強制換新頁**（`page_break_before = True`） |
| 標題 2 以下 | 不強制換頁，隨文流排版 |

實作方式（**兩處均須實作，缺一不可**）：

1. **樣式層級**（初始化文件時，設定 `Heading 1` 樣式）：
   ```python
   s = doc.styles['Heading 1']
   s.paragraph_format.page_break_before = True
   ```

2. **函式層級**（`heading()` 函式內，`level == 1` 時設定）：
   ```python
   if level == 1:
       p.paragraph_format.page_break_before = True
   ```

> **重要**：兩處均須實作，以確保各章節在 Word 開啟時各佔新頁，不得省略。

---

## 一、必要套件

```python
python-docx >= 1.2.0
Pillow (PIL)
fontTools
pypdf  # 用於分析 PDF 監造計畫字型
```

---

## 二、腳本標準結構

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, io, zipfile, re, copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageDraw, ImageFont
from fontTools.ttLib import TTFont as _TTFont

# ─── PIL 字體路徑（sandbox 環境固定路徑）───
FONT_CJK   = "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"
FONT_ASCII = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
```

---

## 三、字型設定（run 層級）

每個 `run` 必須同時設定中文字型和英數字型，**全部使用常數，不寫死**：

```python
def set_run_font(run, size_pt, bold=False, color=None):
    """統一設定 run 字型（使用全域常數）"""
    run.font.name = DOCX_FONT_LATIN
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    # 中文字型設定（必須）
    if run._element.rPr is None:
        run._element.get_or_add_rPr()
    run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    run._element.rPr.rFonts.set(qn('w:ascii'),    DOCX_FONT_LATIN)
    run._element.rPr.rFonts.set(qn('w:hAnsi'),    DOCX_FONT_LATIN)
```

### 樣式層級設定（初始化文件時）

```python
doc = Document()

# Normal 樣式
normal_style = doc.styles['Normal']
normal_style.font.name = DOCX_FONT_LATIN
normal_style.font.size = Pt(DOCX_SZ_BODY)
_nr = normal_style.element.get_or_add_rPr()
_nf = _nr.get_or_add_rFonts()
_nf.set(qn('w:eastAsia'), DOCX_FONT_CJK)
_nf.set(qn('w:ascii'),    DOCX_FONT_LATIN)
_nf.set(qn('w:hAnsi'),    DOCX_FONT_LATIN)

# Heading 樣式（全部黑色，禁用藍色或其他彩色）
HEADING_COLORS = {1: (0,0,0), 2: (0,0,0), 3: (0,0,0)}
for lvl, sz, bold in [(1, DOCX_SZ_H1, True), (2, DOCX_SZ_H2, True), (3, DOCX_SZ_H3, True)]:
    s = doc.styles[f'Heading {lvl}']
    s.font.name = DOCX_FONT_LATIN
    s.font.size = Pt(sz)
    s.font.bold = bold
    s.font.color.rgb = RGBColor(*HEADING_COLORS[lvl])
    s.paragraph_format.space_before = Pt(12)
    s.paragraph_format.space_after  = Pt(6)
    # 標題 1：新章開始強制換新頁（見 §零 0-7）
    if lvl == 1:
        s.paragraph_format.page_break_before = True
    _hr = s.element.get_or_add_rPr()
    _hf = _hr.get_or_add_rFonts()
    _hf.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    _hf.set(qn('w:ascii'),    DOCX_FONT_LATIN)
    _hf.set(qn('w:hAnsi'),    DOCX_FONT_LATIN)
```

---

## 四、常用輔助函式

### heading()：新增標題段落

```python
def heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    # 標題 1：新章開始強制換新頁（見 §零 0-7）
    if level == 1:
        p.paragraph_format.page_break_before = True
    for run in p.runs:
        run.font.name = DOCX_FONT_LATIN
        run.font.color.rgb = RGBColor(0, 0, 0)  # 強制黑色
        run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return p
```

### body()：新增內文段落

```python
def body(doc, text, indent=0):
    p = doc.add_paragraph()
    if indent > 0:
        p.paragraph_format.left_indent = Cm(indent)
    run = p.add_run(text)
    run.font.name = DOCX_FONT_LATIN
    run.font.size = Pt(DOCX_SZ_BODY)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return p
```

### make_table()：新增表格

表格格式規定：
- 表格水平置中
- 標題列背景：淺灰 `F2F2F2`，文字黑色加粗
- 資料列背景：白色 `FFFFFF`，無條紋彩色
- 所有框線：黑色（`Table Grid` 樣式預設）

```python
def make_table(doc, headers, rows, col_widths,
               header_bg='F2F2F2', header_fg=(0,0,0),
               font_size=None):
    """
    header_bg: 標題列背景色，預設淺灰 F2F2F2（禁止藍色或其他彩色）
    header_fg: 標題文字色，預設黑色 (0,0,0)
    注意：已移除 stripe 參數，資料列一律白底黑字
    """
    if font_size is None:
        font_size = DOCX_SZ_TABLE
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 表頭（淺灰背景、黑色加粗文字）
    hdr_row = table.rows[0]
    for ci, (h, w) in enumerate(zip(headers, col_widths)):
        cell = hdr_row.cells[ci]
        cell.width = Cm(w)
        set_cell_bg(cell, header_bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(*header_fg)
        run.font.size = Pt(font_size)
        run.font.name = DOCX_FONT_LATIN
        run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    # 資料列（白底、黑色文字、無條紋）
    for ri, row_data in enumerate(rows):
        row = table.rows[ri+1]
        for ci, (cell_txt, w) in enumerate(zip(row_data, col_widths)):
            cell = row.cells[ci]
            cell.width = Cm(w)
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_txt))
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.name = DOCX_FONT_LATIN
            run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return table
```

### make_toc()：建立目錄（標題1+標題2，頁碼靠右）

目錄格式規定：
- 呈現**標題 1（章）** 與 **標題 2（節）**
- 頁碼靠右對齊（右側 Tab 定位點 + 點線引導）
- 字型大小用 `DOCX_SZ_TOC`；章標題加粗，節標題一般粗細
- 所有文字黑色，禁止彩色

```python
def _add_right_tab(paragraph, pos_cm=15.0):
    """在段落加入靠右定位點（點線引導）"""
    pPr = paragraph._p.get_or_add_pPr()
    # 移除舊的 tabs 元素（避免重複）
    for old in pPr.findall(qn('w:tabs')):
        pPr.remove(old)
    tabs_el = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'),    'right')
    tab.set(qn('w:leader'), 'dot')
    tab.set(qn('w:pos'),    str(int(pos_cm * 567)))  # cm → twips（1cm=567）
    tabs_el.append(tab)
    pPr.append(tabs_el)

def make_toc(doc, entries, right_pos_cm=15.0):
    """
    建立靜態目錄段落。
    entries: list of dicts
      {'level': 1 or 2, 'title': '章節名稱', 'page': 5}
      level=1 → 標題1（章，加粗），level=2 → 標題2（節）

    格式範例：
      壹、計畫範圍 ················· 1
        一、工程概述 ··············· 1

    注意：page 為預估頁碼，開啟 Word 後可手動修正。
    """
    for entry in entries:
        p = doc.add_paragraph()
        _add_right_tab(p, right_pos_cm)

        # 節次縮排
        if entry['level'] == 2:
            p.paragraph_format.left_indent  = Cm(1.0)
            p.paragraph_format.first_line_indent = Cm(0)

        # 章節名稱
        run_title = p.add_run(entry['title'])
        run_title.bold        = (entry['level'] == 1)
        run_title.font.size   = Pt(DOCX_SZ_TOC)
        run_title.font.name   = DOCX_FONT_LATIN
        run_title.font.color.rgb = RGBColor(0, 0, 0)
        run_title._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)

        # Tab + 頁碼（靠右）
        run_page = p.add_run(f'\t{entry["page"]}')
        run_page.bold        = (entry['level'] == 1)
        run_page.font.size   = Pt(DOCX_SZ_TOC)
        run_page.font.name   = DOCX_FONT_LATIN
        run_page.font.color.rgb = RGBColor(0, 0, 0)
        run_page._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return doc
```

### 使用範例

```python
toc_entries = [
    {'level': 1, 'title': '壹、計畫範圍',          'page': 1},
    {'level': 2, 'title': '  一、工程概述',          'page': 1},
    {'level': 2, 'title': '  二、名詞定義',          'page': 3},
    {'level': 1, 'title': '貳、管理權責及分工',       'page': 5},
    {'level': 2, 'title': '  一、品管組織架構',       'page': 5},
    {'level': 2, 'title': '  二、工作職掌',          'page': 7},
    # ... 以此類推
]
make_toc(doc, toc_entries, right_pos_cm=15.0)
```

---

### set_cell_bg()：設定儲存格背景色（含 Bug 修正）

```python
def set_cell_bg(cell, hex_color):
    """設定儲存格背景色，含 python-docx w:shd 位置 bug 修正"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    # Bug 修正：w:shd 必須在 noWrap/vAlign/hideMark 之前
    illegal_tags = {'w:noWrap', 'w:tcMar', 'w:vAlign', 'w:hideMark'}
    first_illegal = next(
        (child for child in tcPr if child.tag.split('}')[-1] in
         {t.split(':')[-1] for t in illegal_tags}), None)
    if first_illegal is not None:
        tcPr.insertBefore(shd, first_illegal)
    else:
        tcPr.append(shd)
```

### add_img()：將 PIL 圖片嵌入文件

```python
def add_img(doc, img_bytes, width_cm=14):
    """將 BytesIO PIL 圖片插入文件"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_bytes, width=Cm(width_cm))
    return p
```

---

## 五、PIL 雙字體方案（v8：僅用於非流程圖的圖片繪製）

> ⚠️ **v8 變更（2026-04-19）**：流程圖**不再使用 PIL PNG**，改用 DrawML
> 向量圖形（見 `08_drawml_integration.md`）。本節保留供其他 PIL 圖片
> 用途（如徽章、浮水印、照片標註）使用。

DroidSansFallbackFull 缺少 ASCII/數學符號，必須雙字體：

```python
FONT_CJK   = "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"
FONT_ASCII = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

# 建立 cmap 索引
_cmap_cjk   = _TTFont(FONT_CJK).getBestCmap()
_cmap_ascii = _TTFont(FONT_ASCII).getBestCmap()

def get_font(size):
    return ImageFont.truetype(FONT_CJK, size)

def get_ascii_font(size):
    return ImageFont.truetype(FONT_ASCII, size)

def _char_font(ch, cjk_f, ascii_f):
    cp = ord(ch)
    if cp in _cmap_cjk: return cjk_f
    if cp in _cmap_ascii: return ascii_f
    return cjk_f  # fallback

def draw_text_mixed(draw, text, xy, cjk_f, ascii_f, fill=(0,0,0)):
    """逐字選字體渲染，解決方框問題"""
    x, y = xy
    for ch in text:
        f = _char_font(ch, cjk_f, ascii_f)
        bb = f.getbbox(ch)
        draw.text((x, y), ch, font=f, fill=fill)
        x += bb[2] - bb[0]

def dtext(draw, xy, text, font, fill=(0,0,0)):
    """便捷函式：混合文字繪製"""
    size = font.size if hasattr(font, 'size') else 13
    af = get_ascii_font(size)
    draw_text_mixed(draw, text, xy, font, af, fill=fill)

def tw(text, font):
    """計算混合文字寬度"""
    size = font.size if hasattr(font, 'size') else 13
    af = get_ascii_font(size)
    w = 0
    for ch in text:
        f = _char_font(ch, font, af)
        bb = f.getbbox(ch)
        w += bb[2] - bb[0]
    return w

def img_to_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf
```

**DroidSansFallbackFull 覆蓋**：中文、全形符號（（）【】＊－＋，。）
**DejaVuSans 覆蓋**：ASCII（-1%）、數學符號（±≦≧→℃°×）、幾何（□○●）

---

## 六、python-docx 已知 Bug 修正

### Bug 1：w:shd 元素位置錯誤（儲存格背景色）

python-docx 將 `w:shd` 放在 `tcPr` 末尾，但 Word 要求在 `noWrap/vAlign/hideMark` 之前。
**已整合在 `set_cell_bg()` 函式中**（見上方），使用 `insertBefore` 修正。

### Bug 2：w:zoom 屬性錯誤（文件縮放）

python-docx 生成 `<w:zoom w:val="bestFit"/>` 而非有效屬性。
必須存檔後用 zipfile 修正：

```python
def fix_docx_bugs(out_path):
    """修正 python-docx 已知 bug（zoom），存檔後執行"""
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(out_path, 'r') as zi, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == 'word/settings.xml':
                data = re.sub(b'<w:zoom[^/]*/>', b'<w:zoom w:percent="100"/>', data)
            zo.writestr(item, data)
    os.replace(tmp, out_path)  # 用 os.replace，避免 PermissionError
    print(f"✅ DOCX 已儲存並修正：{out_path}")
```

---

## 七、頁尾與節區實作（三段式分節）

文件分三個 Section 以實現不同頁尾格式：
- 封面 Section → 無頁尾
- 目錄 Section → 羅馬數字小寫（i, ii, iii…）
- 內文 Section → 阿拉伯數字（1, 2, 3…）

```python
def _add_page_num_field(paragraph):
    """在段落末尾插入 PAGE 域碼（顯示當前頁碼）"""
    run = paragraph.add_run()
    fc_begin = OxmlElement('w:fldChar')
    fc_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    run._r.extend([fc_begin, instr, fc_sep, fc_end])
    return paragraph

def _set_pg_num_type(sectPr, fmt, start=1):
    """設定節區頁碼格式（lowerRoman / decimal）"""
    for old in sectPr.findall(qn('w:pgNumType')):
        sectPr.remove(old)
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:fmt'),  fmt)
    pgNumType.set(qn('w:start'), str(start))
    sectPr.append(pgNumType)

def setup_section_footer(section, fmt):
    """
    設定節區頁尾。
    fmt: 'none' | 'lowerRoman' | 'decimal'
    - 'none'        → 封面：無頁尾
    - 'lowerRoman'  → 目錄：i, ii, iii…
    - 'decimal'     → 內文：1, 2, 3…
    """
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        for r in p.runs:
            r.text = ''
    if fmt == 'none':
        return  # 封面無頁尾，直接返回
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_num_field(p)
    _set_pg_num_type(section._sectPr, fmt, start=1)

def insert_section_break(doc):
    """
    在文件目前位置插入「下一頁」節區分隔符號。
    呼叫後 doc.sections[-1] 即為新 Section。
    """
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    sectPr = OxmlElement('w:sectPr')
    pgSz = OxmlElement('w:pgSz')
    pgSz.set(qn('w:w'), '11906')  # A4 寬（twip）
    pgSz.set(qn('w:h'), '16838')  # A4 高（twip）
    sectPr.append(pgSz)
    pgType = OxmlElement('w:type')
    pgType.set(qn('w:val'), 'nextPage')
    sectPr.append(pgType)
    pPr.append(sectPr)
    return p
```

### 標準三段式頁尾設定流程

```python
# ── 建立文件 ──
doc = Document()

# ── 封面內容 ──
# ... 加封面段落 ...
insert_section_break(doc)   # ← 封面結束

# ── 目錄內容 ──
# ... 加目錄段落 ...
insert_section_break(doc)   # ← 目錄結束

# ── 內文章節 ──
# ... 加各章內容 ...

# ── 設定三個節區的頁尾 ──
sections = doc.sections
setup_section_footer(sections[0], 'none')         # 封面：無頁尾
setup_section_footer(sections[1], 'lowerRoman')   # 目錄：i, ii, iii…
setup_section_footer(sections[2], 'decimal')      # 內文：1, 2, 3…
```

---

## 八、表名與圖名輔助函式

### add_table_caption()：表格上方標題

表名格式：`表{章號}-{章內序號} 表名`，水平置中，位於表格**上方**。

```python
def add_table_caption(doc, chapter, seq, name):
    """
    新增表格標題段落（需在 make_table() 之前呼叫）。
    chapter : 章號（整數，壹=1、貳=2、参=3…）
    seq     : 該章第幾個表（整數，從 1 起算）
    name    : 表格名稱（不含編號）

    範例：
        add_table_caption(doc, chapter=1, seq=1, name="品質管理人員組織表")
        → 產生「表1-1 品質管理人員組織表」
    """
    caption_text = f"表{chapter}-{seq} {name}"
    p = doc.add_paragraph(caption_text)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold = True
    run.font.size = Pt(DOCX_SZ_TABLE)
    run.font.name = DOCX_FONT_LATIN
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return p
```

### add_figure_caption()：圖形下方標題

圖名格式：`圖{章號}-{章內序號} 圖名`，水平置中，位於圖形**下方**。

```python
def add_figure_caption(doc, chapter, seq, name):
    """
    新增圖形標題段落（需在 add_img() 之後呼叫）。
    chapter : 章號（整數，壹=1、貳=2、参=3…）
    seq     : 該章第幾個圖（整數，從 1 起算）
    name    : 圖形名稱（不含編號）

    範例：
        add_img(doc, img_bytes, width_cm=14)
        add_figure_caption(doc, chapter=2, seq=1, name="組織架構圖")
        → 產生「圖2-1 組織架構圖」
    """
    caption_text = f"圖{chapter}-{seq} {name}"
    p = doc.add_paragraph(caption_text)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.size = Pt(DOCX_SZ_BODY)
    run.font.name = DOCX_FONT_LATIN
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), DOCX_FONT_CJK)
    return p
```

### 完整使用範例

```python
# ─── 第壹章（chapter=1）第 1 個表 ───
add_table_caption(doc, chapter=1, seq=1, name="工程主要施工項目及數量表")
make_table(doc, headers=[...], rows=[...], col_widths=[...])

# ─── 第貳章（chapter=2）第 1 個圖 ───
add_img(doc, img_bytes, width_cm=14)
add_figure_caption(doc, chapter=2, seq=1, name="品管組織架構圖")

# ─── 第貳章（chapter=2）第 1 個表 ───
add_table_caption(doc, chapter=2, seq=1, name="工作職掌表")
make_table(doc, headers=[...], rows=[...], col_widths=[...])
```

---

## 九、fc' 符號 Python 語法陷阱

混凝土強度 `fc'=280 kgf/cm²` 中含單引號，在 Python 單引號字串內造成 SyntaxError：

```python
# ❌ 錯誤
('混凝土澆置（fc'=280kgf/cm²）', ...)

# ✅ 正確（跳脫單引號）
('混凝土澆置（fc\'=280kgf/cm²）', ...)

# ✅ 或用雙引號包覆
("混凝土澆置（fc'=280kgf/cm²）", ...)
```

---

## 十、三欄式施工作業流程圖（v8：DrawML 版本）

> ⚠️ **v8 變更**：三欄式改為 python-docx 3 欄表格 + 欄 1 內嵌 DrawML 單節點形狀。
> 整體品質計畫全面停用舊版 PIL PNG 繪製的三欄圖。細節見
> `08_drawml_integration.md` §2.2。

```python
from drawml_flowchart import add_three_col_flow

items = [
    {'flow': '放樣', 'check': '位置、尺寸是否符合圖說', 'doc': '放樣紀錄表', 'mark': '※', 'type': 'process'},
    {'flow': '整地', 'check': '高程、坡度、壓實度',      'doc': '整地檢查紀錄',  'type': 'process'},
    ...
]

add_three_col_flow(doc, '圖伍-7-1 植栽工程施工作業流程及檢驗圖', items,
                   marker_prefix='plant', font_cjk=DOCX_FONT_CJK,
                   font_latin=DOCX_FONT_LATIN, font_pt=DOCX_SZ_TABLE)
```

- `mark`: `'※'`（自主查驗停留點）、`'＊'`（監造停留點）、`'☆'`（檢驗停留點）、`'◎'`（安衛查驗點）
- `type`: 節點圖形（`process`/`decision`/`terminator`/`data`/`document`）

---

## 十一、完整腳本收尾（v8：新增 finalize_flowcharts）

```python
# 1. 儲存文件
doc.save(OUT_PATH)

# 2. 替換 DrawML 占位符（v8 新增，務必在 fix_docx 之前執行）
finalize_flowcharts(OUT_PATH, theme_name='quality_plan',
                    font_cjk=DOCX_FONT_CJK, font_latin=DOCX_FONT_LATIN,
                    font_pt=max(DOCX_SZ_TABLE - 1, 9))

# 3. 修正 python-docx 已知 bug（w:zoom 等）
fix_docx(OUT_PATH)

print(f"✅ 品質計畫已產出：{OUT_PATH}")
```
