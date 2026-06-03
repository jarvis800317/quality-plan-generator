# web-chapter-selector 部署與技術參考

## 專案位置

```
本地開發：~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/web-chapter-selector/
GitHub：https://github.com/jarvis800317/web-chapter-selector
GitHub Pages：https://jarvis800317.github.io/web-chapter-selector/
```

## 工具定位（重要修正，2026-06-03）

這是**常駐工具 URL**，非臨時測試用。主任技師隨時可以打開這個 URL 操作章節篩選，輸出結果直接複製貼入品質計畫文件。

## 部署方式：orphan gh-pages branch（已實作，2026-06-03）

`dist/` 在 `.gitignore` 裡，無法透過一般 git push 部署。採用 orphan `gh-pages` branch 策略。

**完整部署流程（每次修改 source code 後執行）**：

```bash
# 1. 本地 build
cd ~/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/web-chapter-selector
npm run build

# 2. 上傳 dist 到 gh-pages（Python script）
python3 /tmp/deploy_pages_clean.py

# 3. 觸發 GitHub Pages 重建（清除 CDN cache）
gh api repos/jarvis800317/web-chapter-selector/pages/builds -X POST
```

**`/tmp/deploy_pages_clean.py` 核心邏輯**：
- 讀取 dist/ 所有檔案 → base64 編碼 → 建立 Git blob → 建立 tree → 建立 orphan commit → force-push 到 `gh-pages`
- **關鍵：建立 blob 時用 `--input` 吃暫存檔（不用 stdin）**
- GitHub Pages 設定為 `legacy` 模式，從 `gh-pages` branch 的 `/` 路徑讀取

**Token scope 限制**：若 gh auth token 沒有 `workflow` scope，無法 `git push` workflow 檔案。此時 workflow deploy.yml 無法自動部署，只能用上面這套 Python script 流程手動上傳。若要全自動化，須在 GitHub Developer settings 更新 token 加入 `workflow` scope。

## ⚠️ 常見部署失敗：dist/ 在 .gitignore 裡

**根本問題**：`dist/`（編譯產物）被 `.gitignore` 忽略，所以從未 push 到 GitHub。GitHub Pages 只好把原始 `.tsx` 拿出來 serving，瀏覽器無法執行，畫面空白。

**徵狀**：`curl https://jarvis800317.github.io/web-chapter-selector/assets/index-*.js` 全部 404，但 `index.html` 可以存取（內容是 TSX 原始碼）。

**修復**：用上面「部署方式」一節的 Python script 流程。

## ⚠️ GitHub Pages 子目錄部署陷阱（Vite assets 404）

**問題**：Vite 預設用絕對路徑 `/assets/index-*.js`，但 GitHub Pages 的 URL 是 `/web-chapter-selector/assets/...`，導致所有 JS/CSS 返回 404。

**解法**：`vite.config.ts` 必須設 `base: './'`

```ts
export default defineConfig({
  plugins: [react()],
  base: './',   // ← 這行是關鍵
})
```

**驗證**：部署後檢查 `dist/index.html` 內的 `src` 是否為 `./assets/...`（相對路徑），而不是 `/assets/...`（絕對路徑）。

## 按鈕與事件對照表

| 按鈕 | 行為 |
|------|------|
| 下載章節 (.txt) | 觸發瀏覽器下載對話框（本機 .txt 檔案）|
| ✓ 確認並產生品質計畫 | `window.dispatchEvent(CustomEvent('quality-plan-trigger', { detail: {...} }))` — 廣播事件，攜帶 `projectName`/`amount`/`selectedIds`/`timestamp` |

## CDP 端監聽（待實作）

CDP 端監聽 `quality-plan-trigger` 事件的實作方向：

```python
# 在 CDP Python 腳本中，用 page.evaluate 訂閱 window 事件
script = """
window.addEventListener('quality-plan-trigger', (e) => {
  window._hermes_trigger_data = e.detail;
});
"""
page.evaluate(script)

# 之後每輪 CDP 輪詢，檢查是否有新事件
result = page.evaluate("return window._hermes_trigger_data")
if result:
    project_name = result['projectName']
    amount = result['amount']
    selected_ids = result['selectedIds']
    # 啟動 Python 協調腳本
```

完整自動化的其餘部分（Python 協調腳本 + DOCX 產出）待實作。