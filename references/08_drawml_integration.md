# DrawML 流程圖整合規格（v8 強制必讀）

> 本文件說明 `quality-plan-maker` 如何與 `drawml_flowchart.py` 整合，**所有**
> 整體品質計畫內的流程圖（材料送審、施工檢驗、不合格管制、矯正、收發文、
> 設備功能運轉、品質稽核、三欄式施工作業流程及檢驗圖、品管組織架構圖…）
> 一律以 **DrawML**（Office Open XML 向量圖形）繪製，**禁止再用 PIL PNG**。

---

## 零、為什麼改用 DrawML？

| 向度 | PIL PNG（舊） | DrawML（v8） |
|-----|------|------|
| 可編輯性 | ❌ 靜態圖片，Word 內無法改 | ✅ 點擊選取 → 拖曳、縮放、改色、雙擊改字 |
| 向量 / 縮放 | ❌ 放大糊掉 | ✅ 向量；列印、縮放皆清晰 |
| 檔案大小 | 每張圖 ~30–200 KB | 每張 <5 KB |
| 符合金質獎格式規定 | 需要 PIL 自行實作「全黑、無填色」 | `quality_plan` 主題內建 |
| 未來維護 | 需調整 PIL 參數 | 同一 JSON 結構即可產 |

---

## 一、模組位置與相依

| 檔案 | 位置 |
|-----|------|
| `drawml_flowchart.py` | `scripts/drawml_flowchart.py`（Skill 內） |
| 執行期副本 | `mnt/QualityPlanMaker/03_產出品質計畫/drawml_flowchart.py` |

**強制規定**：每次執行 Step 6 前，Claude 必須確認 `drawml_flowchart.py`
已複製到 `03_產出品質計畫/`（與主腳本同目錄），否則 `from drawml_flowchart ...`
會失敗。

相依套件：`lxml`、`python-docx`（sandbox 已預裝）。

---

## 二、三個核心 API

### 2.1 `add_flowchart_placeholder(doc, marker_id, flow_def)`

在 `doc` 目前位置插入一個占位段落，並登錄 `flow_def` 等待後處理。

```python
from drawml_flowchart import add_flowchart_placeholder, preset_material_approval

add_flowchart_placeholder(doc, 'fig_5_1', preset_material_approval())
```

### 2.2 `add_three_col_flow(doc, title, items, marker_prefix='tc', ...)`

建立三欄式表格（作業流程圖｜檢驗要點｜相關記錄/文件），欄 1 為 DrawML
單節點形狀，欄 2/3 為文字儲存格。

```python
items = [
    {'flow': '放樣', 'check': '位置、尺寸符合圖說', 'doc': '放樣紀錄', 'mark': '※', 'type': 'process'},
    {'flow': '整地', 'check': '高程、壓實度',       'doc': '整地紀錄', 'type': 'process'},
    ...
]
add_three_col_flow(doc, '植栽工程施工作業流程及檢驗圖', items, marker_prefix='plant')
```

### 2.3 `finalize_flowcharts(docx_path, theme_name='quality_plan', ...)`

後處理：掃描文件內所有 `<<<DRAWML_FLOWCHART:id>>>` 標記段落，以實際 DrawML
流程圖替換之，並補上 `wpc`/`wps` 命名空間。

```python
doc.save(OUT_PATH)
finalize_flowcharts(OUT_PATH,
                    theme_name='quality_plan',
                    font_cjk=DOCX_FONT_CJK,
                    font_latin=DOCX_FONT_LATIN,
                    font_pt=10)
fix_docx(OUT_PATH)   # 必須在 finalize 之後才做
```

---

## 三、flow_def JSON 結構

```python
flow_def = {
    'nodes': [
        {'id': 'start',  'label': '開始',      'type': 'terminator'},
        {'id': 'input',  'label': '收到申請',  'type': 'data'},
        {'id': 'proc',   'label': '資料驗證',  'type': 'process'},
        {'id': 'judge',  'label': '審核通過？', 'type': 'decision'},
        {'id': 'ok',     'label': '核准通知',   'type': 'process'},
        {'id': 'ng',     'label': '退回修改',   'type': 'process'},
        {'id': 'end',    'label': '結案',       'type': 'terminator'},
    ],
    'edges': [
        {'source': 'start',  'target': 'input'},
        {'source': 'input',  'target': 'proc'},
        {'source': 'proc',   'target': 'judge'},
        {'source': 'judge',  'target': 'ok',  'label': '是'},
        {'source': 'judge',  'target': 'ng',  'label': '否'},
        {'source': 'ng',     'target': 'proc'},     # 退回，形成迴圈（OK 會自動處理反向邊）
        {'source': 'ok',     'target': 'end'},
    ],
    'direction': 'TB',    # 可選：'TB'(預設) 上下 / 'LR' 左右
}
```

**節點類型（`type` 欄）**：

| type | DrawML 圖形 | 對應用途 |
|------|------------|---------|
| `terminator` | 膠囊/橢圓 | 開始 / 結束 |
| `process`    | 矩形     | 一般步驟（預設） |
| `decision`   | 菱形     | 判斷 / 分支 |
| `data`       | 平行四邊形 | 資料輸入/輸出 |
| `document`   | 文件形   | 表單 / 報告輸出 |
| `connector`  | 圓形     | 跨頁連接符號 |

**邊的 label**：用於判斷節點的「是 / 否」分支；其他節點的邊通常不需要 label。

**`\n` 支援**：`label` 中使用 `\n` 可換行（如 `'品管人員\n(王大明)'`）。

---

## 四、預設流程圖（`preset_*` 函式）

直接使用即可，無需自行構造 nodes/edges：

| 函式 | 產出 |
|------|------|
| `preset_material_approval()` | 材料/設備選定送審流程圖（8 節點含迴圈） |
| `preset_work_inspection()`   | 施工檢驗流程圖（10 節點，雙檢驗關卡） |
| `preset_nc_material()`       | 材料自主檢查不合格管制流程圖 |
| `preset_nc_work()`           | 施工自主檢查不合格管制流程圖（含重作迴圈） |
| `preset_corrective()`        | 矯正措施流程圖 |
| `preset_incoming_doc()`      | 收文傳送及歸檔流程圖 |
| `preset_outgoing_doc()`      | 發文傳送及歸檔流程圖 |
| `preset_equipment_test()`    | 設備功能運轉檢測程序流程圖（9 章制景觀專用） |
| `preset_audit_flow()`        | 品質稽核流程圖 |

使用範例：

```python
from drawml_flowchart import add_flowchart_placeholder, preset_nc_material

add_figure_caption(doc, chapter='柒', seq=1, name='材料自主檢查不合格管制流程圖')  # 或按實際章次
add_flowchart_placeholder(doc, 'fig_7_1', preset_nc_material())
```

---

## 五、品管組織架構圖（依 CHAPTER_MODE 動態建）

主腳本的 `make_org_chart()` 已改用 DrawML，且依 `CHAPTER_MODE` 自動決定
層級（`landscape9` 多一層機電工程師）。腳本內直接：

```python
add_img(doc, make_org_chart(), width_cm=14)   # add_img 會辨識 sentinel 自動走 DrawML
```

即可。

---

## 六、`add_img()` sentinel 對照表（相容性）

為了不破壞舊版腳本的 body 程式碼，`make_*_flow()` 回傳 sentinel tuple，
`add_img()` 會辨識並呼叫對應 DrawML API：

| sentinel | 實際行為 |
|----------|---------|
| `('__DRAWML_FLOW__', {nodes, edges})` | 呼叫 `add_flowchart_placeholder()` |
| `('__DRAWML_THREECOL__', {title, items})` | 呼叫 `add_three_col_flow()` |
| `BytesIO`（PIL PNG） | 原有 `add_picture()` 行為（供非流程圖如照片用） |

因此以下寫法在 v8 仍可直接使用：

```python
add_img(doc, make_flow_chart('圖伍-1 材料/設備選定送審流程圖',
                              [...]), width_cm=14)
add_img(doc, make_nc_flow('材料'), width_cm=14)
add_img(doc, make_doc_flow('收文'), width_cm=14)
add_img(doc, make_three_col_flow('圖伍-7-1 植栽工程施工作業流程及檢驗圖',
                                  [...]))
add_img(doc, make_org_chart(), width_cm=14)
```

---

## 七、常見錯誤

### 7.1 `No module named 'drawml_flowchart'`

腳本與 `drawml_flowchart.py` 不在同目錄。解法：
```bash
cp /sessions/.../mnt/.claude/skills/quality-plan-maker/scripts/drawml_flowchart.py \
   /sessions/.../mnt/QualityPlanMaker/03_產出品質計畫/
```

### 7.2 DOCX 開啟時 Word 顯示「內容有問題無法讀取」

通常是：
1. `finalize_flowcharts()` 未呼叫 → 文件內仍有 `<<<DRAWML_FLOWCHART:*>>>` 標記
2. flow_def 中 `source`/`target` 的 id 與 `nodes` 的 id 不一致
3. 自行新增的 node 有重複 id

### 7.3 流程圖形狀位置重疊

- DrawML 自動排版（Sugiyama-style 分層），節點數量建議 ≤ 20 個
- 超過 20 個時，建議拆成多張流程圖，或使用者於 Word 內手動拖曳調整
- 複雜分支情境：可改用 `direction='LR'`（橫向）

### 7.4 「標題」找不到

`build_flowchart_paragraph` 不繪製 title；請用 `add_figure_caption()` 在
流程圖**下方**加上「圖X-X　xxx流程圖」。

### 7.5 中文字型顯示異常

`finalize_flowcharts()` 的 `font_cjk` 參數應與文件其他部份一致
（通常是 `DOCX_FONT_CJK`，即 `'標楷體'` 或監造計畫指定之中文字型）。

---

## 八、Step 5B（全新製作）的標準骨架

```python
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from drawml_flowchart import (
    add_flowchart_placeholder, add_three_col_flow,
    finalize_flowcharts,
    preset_material_approval, preset_nc_material, ...
)

doc = Document()
# ... 寫封面、目錄、各章節 ...

# 流程圖範例：
add_figure_caption(doc, chapter='伍', seq=1, name='材料/設備選定送審流程圖')
add_flowchart_placeholder(doc, 'fig_5_1', preset_material_approval())

# 三欄式範例：
items = [{'flow': ..., 'check': ..., 'doc': ..., 'mark': '※', 'type': 'process'}, ...]
add_three_col_flow(doc, '圖伍-7-1 植栽工程施工作業流程及檢驗圖', items,
                   marker_prefix='plant', font_cjk=DOCX_FONT_CJK,
                   font_latin=DOCX_FONT_LATIN, font_pt=DOCX_SZ_TABLE)

# 收尾：
doc.save(OUT_PATH)
finalize_flowcharts(OUT_PATH, theme_name='quality_plan',
                    font_cjk=DOCX_FONT_CJK, font_latin=DOCX_FONT_LATIN,
                    font_pt=max(DOCX_SZ_TABLE - 1, 9))
fix_docx(OUT_PATH)
```

---

## 九、自訂流程圖（當 `preset_*` 不敷使用）

直接構造 `flow_def`：

```python
custom_flow = {
    'nodes': [
        {'id': 'a', 'label': '步驟1', 'type': 'terminator'},
        {'id': 'b', 'label': '步驟2', 'type': 'process'},
        {'id': 'c', 'label': '合格？', 'type': 'decision'},
        {'id': 'd', 'label': '改善',   'type': 'process'},
        {'id': 'e', 'label': '結案',   'type': 'terminator'},
    ],
    'edges': [
        {'source': 'a', 'target': 'b'},
        {'source': 'b', 'target': 'c'},
        {'source': 'c', 'target': 'd', 'label': '否'},
        {'source': 'd', 'target': 'b'},           # 迴圈：回到步驟 2
        {'source': 'c', 'target': 'e', 'label': '是'},
    ],
}
add_flowchart_placeholder(doc, 'my_flow', custom_flow)
```

---

## 十、不得違反的強制規定

1. **禁止**使用 PIL 產生流程圖 PNG；若腳本 import `PIL.ImageDraw` 去畫流程圖，
   視同 v8 規定違反。
2. **禁止**於流程圖形狀上加填色（`quality_plan` 主題使用 `a:noFill`，
   切勿改為 `solidFill`）。
3. **必須**於 `doc.save()` 之後、`fix_docx()` 之前呼叫
   `finalize_flowcharts(OUT_PATH, ...)`。
4. **必須**將 `drawml_flowchart.py` 複製到主腳本同目錄。
5. **圖名**仍以 `add_figure_caption(doc, chapter, seq, name)` 置於流程圖**下方**
   （金質獎格式要求：表名在上、圖名在下）。
