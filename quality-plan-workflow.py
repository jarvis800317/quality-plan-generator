#!/usr/bin/env /usr/bin/python3
"""
quality-plan-workflow.py
========================
品質計畫自動化流程

流程：
  1. 讀取監造計畫 DOCX → 抽出 工程名稱、契約金額（萬元）
  2. 更新 web-chapter-selector/src/config.json
  3. 等待 Vite HMR 熱更新（約 1-2 秒）
  4. 用 CDP 自動填入瀏覽器表單（精準寫入 React state）
  5. 主任技師在瀏覽器操作 → 下載 .txt → 回報「完成」

用法：
  python3 quality-plan-workflow.py
"""

import subprocess
import json
import time
import sys
import os

# ── 設定區 ──────────────────────────────────────────────────
MONITORING_PLAN = "/Users/handsome/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/01-參考監造計畫/桃林鐵路鋪面改善統包工程監造計畫(初稿).docx"
CONFIG_JSON = "/Users/handsome/我的雲端硬碟/Jarvis專案/projects/quality-plan-generator/02-產出品質計畫/web-chapter-selector/src/config.json"
WEB_APP_URL = "http://localhost:5173"


# ── Step 1: 讀取 DOCX ──────────────────────────────────────
def read_monitoring_plan():
    script = f"""
from docx import Document
doc = Document(r"{MONITORING_PLAN}")
project_name = None
budget_wan = None
for p in doc.paragraphs:
    t = p.text.strip()
    # 工程名稱（只取第一個，且不含後續噪聲關鍵字）
    if not project_name and "工程名稱：" in t and "契約" not in t and "工程主辦" not in t:
        raw = t.split("工程名稱：")[1]
        for kw in ["工程主辦", "契約編號", "監造單位", "設計單位"]:
            if kw in raw:
                raw = raw.split(kw)[0]
        project_name = raw.strip()
    # 工程預算（元 → 萬元）
    if not budget_wan and "工程預算：" in t:
        raw = t.split("工程預算：")[1]
        raw = raw.replace("新台幣","").replace("元整","").replace(",","").replace("。","")
        try:
            budget_wan = str(int(raw) / 10000)
        except ValueError:
            budget_wan = raw.strip()
print(f"PROJECT_NAME={{project_name}}")
print(f"BUDGET_WAN={{budget_wan}}")
"""
    result = subprocess.run(["/usr/bin/python3", "-c", script],
                           capture_output=True, text=True)
    if result.returncode != 0:
        print("❌ 讀取監造計畫失敗：", result.stderr)
        sys.exit(1)

    lines = {}
    for l in result.stdout.strip().split("\n"):
        if "=" in l:
            k, v = l.split("=", 1)
            lines[k.strip()] = v.strip()

    project_name = lines.get("PROJECT_NAME", "").strip()
    budget_wan = lines.get("BUDGET_WAN", "").strip()

    if not project_name or not budget_wan:
        print("❌ 無法從監造計畫抽出工程名稱或契約金額")
        print("原始輸出：", result.stdout)
        sys.exit(1)

    return project_name, budget_wan


# ── Step 2: 更新 config.json ─────────────────────────────
def update_config(project_name: str, budget_wan: str):
    config = {
        "projectName": project_name,
        "amountWan": budget_wan
    }
    with open(CONFIG_JSON, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"✅ config.json 已更新")
    print(f"   工程名稱：{project_name}")
    print(f"   契約金額：{budget_wan} 萬元")


# ── 主流程 ─────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 52)
    print("  品質計畫自動化流程")
    print("=" * 52)

    print("\n【Step 1】讀取監造計畫 DOCX...")
    project_name, budget_wan = read_monitoring_plan()

    print("\n【Step 2】更新 web-chapter-selector 設定檔...")
    update_config(project_name, budget_wan)

    print("\n【Step 3】等待 Vite HMR 熱更新（~2 秒）...")
    time.sleep(2)

    print(f"\n【Step 4】啟動瀏覽器 → {WEB_APP_URL}")
    print("   → 系統將自動填入工程名稱與契約金額")
    print("   → 主任技師確認 / 調整章節 → 點「下載章節(.txt)」")
    print("   → 完成後回報「完成」，我繼續後續處理")
    print("=" * 52)