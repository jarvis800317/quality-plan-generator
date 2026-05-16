---
name: quality-plan-generator
description: >
  品質計畫自動產生器。適用情境：主任技師或品管人員要依據「監造計畫」（PDF 或 DOCX）
  自動產生「整體品質計畫」Word 文件。觸發關鍵字：品質計畫、品管計畫、
  整體品質計畫、quality plan、依監造計畫做品質計畫。
  工程類型涵蓋景觀（金質獎 8/9 章）、道路（11 章）、建築、水利，以及尚無範本的工程類型。

  📂 固定工作目錄：
    監造計畫輸入 → ~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/01-參考監造計畫/
    品質計畫輸出 → ~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/
    腳本模板     → ~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/03-腳本模板/
    規格文件     → ~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/references/

linked_files:
  references:
    - references/01_workflow.md
    - references/02_landscape.md
    - references/03_road.md
    - references/04_general.md
    - references/05_correspondence.md
    - references/06_docx_tech.md
    - references/07_chapter_selection.md
    - references/08_drawml_integration.md
---

# 品質計畫自動產生器

## 定位

主任技師上傳工程「監造計畫」後，本 Skill 引導 AI：
1. 分析監造計畫（提取工程資訊、字型規格、施工工項）
2. 自動判斷工程類型與章節制
3. 執行 Python 腳本，產生完整的 DOCX「整體品質計畫」
4. 輸出至 `02-產出品質計畫/` 資料夾（使用絕對路徑）

---

## 標準開場指令

使用者貼上以下指令後，AI 直接執行，**不詢問模式選擇**：

```
我是營造業主任技師，要製作品質計畫。
監造計畫檔案：quality-plan-generator/01-參考監造計畫/【檔名】
```

---

## AI 執行流程（每次必須完整走完）

### Step 1 — 讀取規格文件（Skill 內建，无需上傳）

依工程類型讀取本 Skill 的 `references/` 目錄：

| 檔案 | 讀取時機 |
|-----|---------|
| `references/01_workflow.md` | 每次都讀（完整作業流程與規則） |
| `references/02_landscape.md` | 工程類型為景觀/公園/市地重劃 |
| `references/03_road.md` | 工程類型為道路/排水/箱涵 |
| `references/04_general.md` | 建築/水利/其他新類型工程 |
| `references/05_correspondence.md` | 需要章節對應關係時 |
| `references/06_docx_tech.md` | 撰寫/修改 Python 腳本時 |
| `references/07_chapter_selection.md` | **每次都讀**（Step 4 章節選擇） |
| `references/08_drawml_integration.md` | **每次都讀**（流程圖用 DrawML） |

> ⚠️ **規格文件已內建於本 Skill**，不需要從 GDrive 讀取。
>
> ⚠️ **雙版本架構（重要）**：本 Skill 的 `references/` 下的規格檔是**實際執行時使用的唯一正確版本（Source of Truth）**。GDrive 的 `quality-plan-generator/04-規格文件/` 僅供人工參閱。**你在 GDrive 版編輯後，必須同步更新本 Skill 的 `references/` 檔案**，否則變更不會生效。兩個版本可能不同步，GDrive 版通常較新。
>
> ⚠️ **08_drawml_integration.md 尚不存在**，流程圖技術規格目前暫用 `02_landscape.md` §五與 `06_docx_tech.md` 的 DrawML 整合說明。
>
> ⚠️ **雙版本警告（重要）**：`references/` 下的規格檔是實際執行時使用的規則，是**唯一正確版本**。GDrive 的 `quality-plan-generator/04-規格文件/` 僅供人工參閱。**編輯 GDrive 版後，必須同步更新本 Skill 的 `references/` 檔案**，否則變更不會生效。本 Skill 的 `references/` 是事實來源（source of truth）。

### Step 2 — 分析監造計畫

讀取 `01-參考監造計畫/` 內的監造計畫檔案，提取：

**A. 字型規格（每次必分析，不得沿用舊值）**
- 中文字型名稱（如：標楷體、新細明體）
- 英數字型名稱（如：Arial、Times New Roman）
- 各層級字型大小

**B. 工程基本資訊**
- 工程名稱、主辦機關、設計/監造單位、承攬廠商、工地主任、品管人員、主任技師
- 契約金額、工地地點、標案名稱

**C. 主要施工工項及數量**
- 直接從監造計畫表格完整複製（見 `01_workflow.md` §十一，嚴禁截斷）

**D. 判斷章節制**

| 工程特徵 | 章節制 |
|---------|-------|
| 景觀/公園，有機電設備（噴灌/照明/排水設備） | 景觀 9 章（壹～玖） |
| 景觀/公園，無機電設備 | 景觀 8 章（壹～捌） |
| 道路/排水/箱涵 | 道路 11 章（第一～第十一章） |
| 建築/水利/其他，金額 ≥5,000 萬 | 一般 11 章 |
| 建築/水利/其他，金額 1,000 萬～5,000 萬 | 一般 6 章 |
| 建築/水利/其他，金額 150 萬～1,000 萬 | 一般 4 章（法定必選） |

### Step 3 — 顯示工程概要確認表（10 行以內）

```
【工程概要確認】
工程名稱：＿＿＿
主辦機關：＿＿＿
監造單位：＿＿＿
承攬廠商：＿＿＿（請補充）
契約金額：＿＿＿元
章節制：＿＿＿（景觀9章/景觀8章/道路11章/一般11章/一般6章/一般4章）
字型：中文＿＿＿、英數＿＿＿
內文＿＿pt、章標題＿＿pt、節標題＿＿pt
主要工項：＿＿項
```

等主任技師確認（或補充資訊）後，繼續執行 Step 4。

---

### Step 4 — 章節選擇（強制必執行，不得跳過）

依據 `07_chapter_selection.md` 的邏輯，用 `clarify` 工具以選擇題方式讓主任技師確認章節配置：

**使用 clarify 選擇題**，四個選項如下（直接給出，不要讓用戶輸入）：

- **選項 A（11章制）**：全 11 章（第一～第十一章）— 5,000 萬以上工程
- **選項 B（景觀9章）**：景觀 9 章（壹～玖）— 有機電設備
- **選項 C（景觀8章）**：景觀 8 章（壹～捌）— 無機電設備
- **選項 D（金額分級）**：依契約金額自動判斷（150萬以下→無強制/150-1000萬→4章/1000-5000萬→6章/5000萬以上→11章）

確認後進入 Step 5。

---

### Step 5 — 選擇腳本模式

**A. 有同類型 .py 範本時（腳本複用模式）**
1. 讀取 `~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/03-腳本模板/quality_plan_template.py`
2. 修改設定區常數（工程資訊 + 字型設定）
3. 執行腳本，產生 DOCX

**B. 無同類型 .py 範本時（全新製作模式）**
1. 讀取 `references/06_docx_tech.md` 取得技術規格
2. 依章節制從對應規格檔讀取每章規格
3. 將腳本寫入 `02-產出品質計畫/`，執行產生 DOCX

---

### Step 6 — 執行 Python 腳本

⚠️ **路徑強制規定**：`OUT_PATH` 必須使用從 `__file__` 動態計算的絕對路徑。

腳本內的路徑計算（不得改用相對路徑）：
```python
import os
# 腳本位於 quality-plan-generator/02-產出品質計畫/<script>.py
_WORKSPACE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(_WORKSPACE, f"{ENG_NAME}-整體品質計畫.docx")
```

執行範例：
```bash
python3 /Users/handsome/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/quality_plan_景觀.py
```

⚠️ **DrawML 流程圖強制規定**：
- 腳本必須與 `drawml_flowchart.py` 同目錄（`02-產出品質計畫/`）
- `doc.save(OUT_PATH)` 之後、`fix_docx(OUT_PATH)` 之前呼叫 `finalize_flowcharts(OUT_PATH, theme_name='quality_plan', ...)`

---

### Step 7 — 收尾

1. 將腳本存於 `02-產出品質計畫/quality_plan_[類型].py`（供下次複用）
2. 確認 `drawml_flowchart.py` 已在同目錄
3. 將 DOCX 交付給使用者（路徑：`~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/`）

---

## 強制規定

### 流程規定

1. **Step 4 章節選擇為強制關卡**：使用 `clarify` 工具選擇題確認
2. **流程圖一律用 DrawML**（向量化可編輯），使用 `add_flowchart_placeholder()` + `finalize_flowcharts()`；主題一律為 `quality_plan`（無填色＋黑色框線＋黑色文字）
3. **不在對話中完整展示章節文字**，全部直接寫入 Python 腳本
4. **不逐章詢問「下一章」**，全自動跑完
5. **執行完後修正 python-docx 已知 Bug**（見 `06_docx_tech.md` §六）

### 格式規定（撰寫腳本時必須逐條實作）

6. **頁尾三段式**：封面無頁尾、目錄羅馬數字（i/ii/iii…）、內文阿拉伯數字（1/2/3…）
7. **目錄格式**：標題 1（章）與標題 2（節），頁碼靠右對齊
8. **全文黑色**：所有文字/標題/表格框線/圖形框線均為 RGB(0,0,0)；圖形無顏色填滿
9. **字型彈性**：每次依監造計畫分析結果設定，全部使用常數 `DOCX_FONT_CJK`、`DOCX_FONT_LATIN`、`DOCX_SZ_*`
10. **表格編號**：表名在**上方**置中，格式 `表{章號}-{序號} 表名`
11. **圖形編號**：圖名在**下方**置中，格式 `圖{章號}-{序號} 圖名`
12. **「表1-1 工程主要施工項目及數量表」必須完整複製**：禁止截斷、刪減、合併
13. **材料設備送審管制表（表伍-28）**：14 欄格式（編號/審查項目/提出時機/提出時限/承攬商應備文件/監造單位審查意見/核定/備註）
14. **三欄式流程圖**：作業流程圖（含不合格退回路徑）｜ 檢驗要點 ｜ 相關記錄/文件

---

## 主詞轉換規則（監造→品質）

| 監造計畫主詞 | 品質計畫主詞 |
|------------|------------|
| 監造單位 | 本公司（承攬廠商） |
| 本公司審查 | 提送監造單位審查 |
| 通知廠商改善 | 自行辦理改善 |
| 監造人員簽認 | 工地主任/現場施工人員簽認 |

---

## 停留點符號規定

| 工程類型 | 符號 | 意義 |
|---------|-----|------|
| 景觀（金質獎） | ＊ | 監造停留點 |
| 景觀（金質獎） | ※ | 自主檢查停留點 |
| 道路（金質獎） | ☆ | 自主檢查停留點 |
| 道路（金質獎） | ◎ | 安衛查驗點 |
| 一般版 | ▽ | 檢驗停留點 |

---

## 依賴套件

```bash
pip install python-docx pillow fonttools pypdf pdfplumber lxml
```

第一次使用時，AI 會提醒用戶確認這些套件已安裝。