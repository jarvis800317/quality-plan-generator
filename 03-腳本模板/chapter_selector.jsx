import { useState, useCallback, useRef } from "react";

// ═══════════════════════════════════════════════════════════
// 品質計畫 — 章節選擇器  (Quality Plan Chapter Selector)
// Step 4 ─ 完整11章彈性勾選版（金額單位：元）
// ═══════════════════════════════════════════════════════════

const ALL_CHAPTERS = [
  { key: "scope",      zhNum: "壹",   title: "計畫範圍",            note: "" },
  { key: "mgmt",       zhNum: "貳",   title: "管理責任及權責分工",   note: "" },
  { key: "method",     zhNum: "參",   title: "施工要領",            note: "" },
  { key: "standard",   zhNum: "肆",   title: "品質管理標準",         note: "" },
  { key: "inspection", zhNum: "伍",   title: "材料及施工檢驗程序",   note: "" },
  { key: "equipment",  zhNum: "陸",   title: "設備功能運轉測試",     note: "有機電設備時建議保留" },
  { key: "selfcheck",  zhNum: "柒",   title: "自主檢查表",           note: "" },
  { key: "nc",         zhNum: "捌",   title: "不合格品之管制",       note: "" },
  { key: "ca",         zhNum: "玖",   title: "矯正與預防措施",       note: "" },
  { key: "audit",      zhNum: "拾",   title: "內部品質稽核",         note: "" },
  { key: "doc",        zhNum: "拾壹", title: "文件紀錄管理系統",     note: "" },
];

// 金額單位：元
// 5,000萬 = 50,000,000 元
// 1,000萬 = 10,000,000 元
//   150萬 =  1,500,000 元
function getMandatoryKeys(amount) {
  if (amount >= 50000000) {
    return ["scope","mgmt","method","standard","inspection","selfcheck","nc","ca","audit","doc"];
  } else if (amount >= 10000000) {
    return ["scope","mgmt","standard","inspection","selfcheck","doc"];
  } else if (amount >= 1500000) {
    return ["mgmt","standard","inspection","selfcheck"];
  }
  return [];
}

function getDefaultSelected(amount) {
  if (amount >= 50000000) {
    return ALL_CHAPTERS.map(c => c.key);
  } else if (amount >= 10000000) {
    return ["scope","mgmt","method","standard","inspection","selfcheck","doc"];
  } else if (amount >= 1500000) {
    return ["scope","mgmt","standard","inspection","selfcheck","doc"];
  }
  return ALL_CHAPTERS.map(c => c.key);
}

function getLevelInfo(amount) {
  if (amount >= 50000000)  return { label: "5,000萬以上",          color: "#1a56db", bg: "#ebf5ff", mandatoryCount: 10 };
  if (amount >= 10000000)  return { label: "1,000萬～未達5,000萬", color: "#0e9f6e", bg: "#f0fdf4", mandatoryCount: 6  };
  if (amount >= 1500000)   return { label: "150萬～未達1,000萬",   color: "#d97706", bg: "#fffbeb", mandatoryCount: 4  };
  return { label: "未達150萬", color: "#9ca3af", bg: "#f9fafb", mandatoryCount: 0 };
}

// 數字格式化（千分位）
function formatAmount(val) {
  if (!val) return "";
  return Number(val).toLocaleString("zh-TW");
}

const zhNums = ["壹","貳","參","肆","伍","陸","柒","捌","玖","拾","拾壹"];

// ══ 契約金額預設值（Claude 每次產出前會更新此行） ──────────────
// Claude 使用時：將下方數值替換為當次監造計畫分析出的契約金額
const CONTRACT_AMOUNT_FROM_PLAN = 107947000; // ← Claude 更新此值
// ────────────────────────────────────────────────────────────

export default function ChapterSelector() {
  const [amountStr, setAmountStr] = useState(String(CONTRACT_AMOUNT_FROM_PLAN));
  const [confirmed, setConfirmed] = useState(false);
  const [selected, setSelected] = useState(() => {
    const defaults = getDefaultSelected(CONTRACT_AMOUNT_FROM_PLAN);
    const init = {};
    ALL_CHAPTERS.forEach(c => { init[c.key] = defaults.includes(c.key); });
    return init;
  });
  const [copyStatus, setCopyStatus] = useState("idle"); // "idle" | "show" | "copied"
  const textareaRef = useRef(null);

  const amount = parseFloat(amountStr.replace(/,/g, "")) || 0;
  const mandatoryKeys = getMandatoryKeys(amount);
  const levelInfo = getLevelInfo(amount);

  // ★ 必須在 buildCopyText 之前定義
  const selectedChapters = ALL_CHAPTERS.filter(c => selected[c.key]);

  const handleAmountChange = useCallback((val) => {
    // 只允許數字（移除非數字字元後存入）
    const clean = val.replace(/[^0-9]/g, "");
    setAmountStr(clean);
    setConfirmed(false);
    const parsed = parseFloat(clean) || 0;
    const defaults = getDefaultSelected(parsed);
    const next = {};
    ALL_CHAPTERS.forEach(c => { next[c.key] = defaults.includes(c.key); });
    setSelected(next);
  }, []);

  const handleConfirmAmount = () => {
    if (amount >= 1500000) { setConfirmed(true); setCopyStatus("idle"); }
  };

  // 產生給 Claude 解析的結構化複製文字
  const buildCopyText = useCallback(() => {
    const levelInfo = getLevelInfo(amount);
    const included = selectedChapters;
    const excluded = ALL_CHAPTERS.filter(c => !selected[c.key]);
    const flagLine = ALL_CHAPTERS.map(c => `${c.key}:${selected[c.key] ? "T" : "F"}`).join(",");
    const includedStr = included.map((c, i) => `${zhNums[i]} ${c.title}`).join("、");
    const excludedStr = excluded.length > 0 ? excluded.map(c => `${c.zhNum} ${c.title}`).join("、") : "無";
    return [
      "【章節配置確認 — 請將以下文字貼至 Claude 對話框】",
      `契約金額：${amount.toLocaleString("zh-TW")} 元（${levelInfo.label}，法規必選 ${levelInfo.mandatoryCount} 章）`,
      `納入章節（共 ${included.length} 章）：${includedStr}`,
      `排除章節：${excludedStr}`,
      `CHAPTER_FLAGS=${flagLine}`,
    ].join("\n");
  }, [amount, selected, selectedChapters]);

  // 步驟1：顯示文字框；步驟2：自動全選供手動 Ctrl+C
  const handleCopy = useCallback(() => {
    setCopyStatus("show");
    // 下一幀等 textarea render 後再全選
    setTimeout(() => {
      if (textareaRef.current) {
        textareaRef.current.focus();
        textareaRef.current.select();
        // 嘗試 clipboard API（成功則改狀態，失敗維持 show 讓使用者手動 Ctrl+C）
        try {
          navigator.clipboard.writeText(textareaRef.current.value)
            .then(() => setCopyStatus("copied"))
            .catch(() => {}); // 失敗不報錯，使用者已全選可自行 Ctrl+C
        } catch (_) {}
      }
    }, 80);
  }, [buildCopyText]);

  const handleTextareaClick = useCallback(() => {
    if (textareaRef.current) {
      textareaRef.current.select();
    }
  }, []);

  const toggleChapter = (key) => {
    if (mandatoryKeys.includes(key)) return;
    setSelected(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const selectAll = () => {
    const next = {};
    ALL_CHAPTERS.forEach(c => { next[c.key] = true; });
    setSelected(next);
  };

  const selectMandatoryOnly = () => {
    const next = {};
    ALL_CHAPTERS.forEach(c => { next[c.key] = mandatoryKeys.includes(c.key); });
    setSelected(next);
  };

  const card = {
    background: "#fff",
    borderRadius: 12,
    boxShadow: "0 2px 12px rgba(0,0,0,0.09)",
    padding: "22px 26px",
    marginBottom: 16,
  };

  return (
    <div style={{ fontFamily: "'Noto Sans TC','PingFang TC',sans-serif", maxWidth: 640, margin: "0 auto", padding: "24px 16px", color: "#1f2937" }}>

      {/* ── 標題 ── */}
      <div style={{ textAlign: "center", marginBottom: 24 }}>
        <div style={{ fontSize: 22, fontWeight: 800, color: "#1a56db", marginBottom: 4 }}>
          📋 整體品質計畫
        </div>
        <div style={{ fontSize: 13, color: "#6b7280" }}>章節選擇器　Step 4　—　完整11章彈性勾選</div>
      </div>

      {/* ══ Step 1：確認契約金額 ══ */}
      <div style={card}>
        <div style={{ fontWeight: 700, marginBottom: 6, fontSize: 14, color: "#374151" }}>
          ❶　確認契約金額
        </div>
        <div style={{
          fontSize: 12, color: "#6b7280", marginBottom: 12,
          padding: "6px 10px", background: "#f0f9ff", borderRadius: 6,
          borderLeft: "3px solid #38bdf8",
        }}>
          📄 已從監造計畫讀取契約金額，如有誤差請直接修改後確認
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <input
            type="text"
            inputMode="numeric"
            value={amount > 0 ? amount.toLocaleString("zh-TW") : ""}
            onChange={e => handleAmountChange(e.target.value)}
            style={{
              flex: 1, padding: "10px 14px", borderRadius: 8,
              border: "1.5px solid #93c5fd", fontSize: 16,
              outline: "none", fontWeight: 600, color: "#1e40af",
              background: "#f0f6ff", letterSpacing: 1,
            }}
          />
          <span style={{ color: "#6b7280", whiteSpace: "nowrap", fontWeight: 600 }}>元</span>
          <button
            onClick={handleConfirmAmount}
            disabled={amount < 1500000}
            style={{
              padding: "10px 18px", borderRadius: 8, border: "none",
              background: amount >= 1500000 ? "#1a56db" : "#e5e7eb",
              color: amount >= 1500000 ? "#fff" : "#9ca3af",
              fontWeight: 700, cursor: amount >= 1500000 ? "pointer" : "not-allowed",
              fontSize: 14, whiteSpace: "nowrap",
            }}
          >
            確認金額
          </button>
        </div>

        {amount >= 1500000 && (
          <div style={{
            marginTop: 10, padding: "6px 14px", borderRadius: 20,
            background: levelInfo.bg, color: levelInfo.color,
            fontSize: 13, fontWeight: 600, display: "inline-flex",
            alignItems: "center", gap: 6,
            border: `1px solid ${levelInfo.color}44`,
          }}>
            📊 {levelInfo.label}　｜　法規必選 {levelInfo.mandatoryCount} 章
          </div>
        )}
        {amount > 0 && amount < 1500000 && (
          <div style={{ marginTop: 8, color: "#ef4444", fontSize: 13 }}>
            ⚠️ 契約金額未達 1,500,000 元（150萬），請洽監造單位確認適用規範
          </div>
        )}
      </div>

      {/* ══ Step 2：章節勾選（完整11章） ══ */}
      {confirmed && (
        <div style={card}>
          <div style={{ fontWeight: 700, marginBottom: 4, fontSize: 14, color: "#374151" }}>
            ❷　勾選章節
            <span style={{ fontWeight: 400, color: "#6b7280" }}>
              （已選 <strong style={{ color: "#1a56db" }}>{selectedChapters.length}</strong> / 11 章）
            </span>
          </div>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 14 }}>
            🔒 必選章節已鎖定不可取消　｜　➕ / 💡 選配章節可自由勾選
          </div>

          {/* 快速選取按鈕 */}
          <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
            {[
              { label: "✅ 全選 11 章", action: selectAll },
              { label: `🔒 僅必選 ${mandatoryKeys.length} 章`, action: selectMandatoryOnly },
            ].map(btn => (
              <button
                key={btn.label}
                onClick={btn.action}
                style={{
                  padding: "6px 14px", borderRadius: 20,
                  border: "1.5px solid #d1d5db",
                  background: "#f9fafb", color: "#374151",
                  fontSize: 13, cursor: "pointer", fontWeight: 500,
                }}
              >
                {btn.label}
              </button>
            ))}
          </div>

          {/* 完整 11 章逐條勾選清單 */}
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {ALL_CHAPTERS.map((ch) => {
              const isMandatory = mandatoryKeys.includes(ch.key);
              const isSelected  = !!selected[ch.key];
              const isEquip     = ch.key === "equipment";

              return (
                <div
                  key={ch.key}
                  onClick={() => toggleChapter(ch.key)}
                  style={{
                    display: "flex", alignItems: "center", gap: 12,
                    padding: "10px 14px", borderRadius: 8,
                    border: isSelected
                      ? (isMandatory ? "1.5px solid #1a56db" : "1.5px solid #0e9f6e")
                      : "1.5px solid #e5e7eb",
                    background: isSelected
                      ? (isMandatory ? "#ebf5ff" : "#f0fdf4")
                      : "#fafafa",
                    cursor: isMandatory ? "not-allowed" : "pointer",
                    transition: "all 0.15s",
                    opacity: (!isMandatory && !isSelected) ? 0.6 : 1,
                  }}
                >
                  <div style={{
                    width: 20, height: 20, borderRadius: 4, flexShrink: 0,
                    border: isSelected
                      ? (isMandatory ? "2px solid #1a56db" : "2px solid #0e9f6e")
                      : "2px solid #d1d5db",
                    background: isSelected
                      ? (isMandatory ? "#1a56db" : "#0e9f6e")
                      : "#fff",
                    display: "flex", alignItems: "center", justifyContent: "center",
                  }}>
                    {isSelected && <span style={{ color: "#fff", fontSize: 13, lineHeight: 1 }}>✓</span>}
                  </div>

                  <div style={{ flex: 1 }}>
                    <span style={{ fontWeight: 600, fontSize: 14 }}>
                      {ch.zhNum}　{ch.title}
                    </span>
                    {ch.note && (
                      <span style={{ fontSize: 12, color: "#78716c", marginLeft: 8 }}>
                        💡 {ch.note}
                      </span>
                    )}
                  </div>

                  <div style={{
                    fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 10,
                    background: isMandatory ? "#dbeafe" : (isEquip ? "#fef3c7" : "#f3f4f6"),
                    color: isMandatory ? "#1e40af" : (isEquip ? "#92400e" : "#6b7280"),
                    whiteSpace: "nowrap",
                  }}>
                    {isMandatory ? "🔒 必選" : (isEquip ? "💡 選配" : "➕ 選配")}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ══ Step 3：確認章節清單並產出 ══ */}
      {confirmed && (
        <div style={card}>
          <div style={{ fontWeight: 700, marginBottom: 12, fontSize: 14, color: "#374151" }}>
            ❸　確認章節清單（動態重編章序）
          </div>

          <div style={{
            background: "#f8fafc", borderRadius: 8, padding: "12px 16px",
            marginBottom: 14, fontSize: 13.5, lineHeight: 2.1,
            border: "1px dashed #cbd5e1",
          }}>
            {selectedChapters.length === 0
              ? <span style={{ color: "#9ca3af" }}>（尚未選取任何章節）</span>
              : selectedChapters.map((ch, idx) => (
                  <div key={ch.key} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <strong style={{ minWidth: 32, color: "#374151" }}>{zhNums[idx]}</strong>
                    <span>{ch.title}</span>
                    {mandatoryKeys.includes(ch.key)
                      ? <span style={{ fontSize: 10, background: "#dbeafe", color: "#1e40af", padding: "1px 6px", borderRadius: 8 }}>必選</span>
                      : <span style={{ fontSize: 10, background: "#d1fae5", color: "#065f46", padding: "1px 6px", borderRadius: 8 }}>選配</span>
                    }
                  </div>
                ))
            }
          </div>

          <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 16 }}>
            共 <strong style={{ color: "#1f2937", fontSize: 15 }}>{selectedChapters.length}</strong> 章，章序已自動重新編排（不跳號）
          </div>

          {/* 說明文字 */}
          <div style={{
            fontSize: 12, color: "#6b7280", marginBottom: 12,
            padding: "8px 12px", background: "#fffbeb", borderRadius: 6,
            borderLeft: "3px solid #fbbf24",
          }}>
            📌 確認章節後，點「複製配置」，再將文字<strong>貼至 Claude 對話框</strong>，Claude 即可直接讀取並產出文件，無需其他操作。
          </div>

          {/* 複製按鈕 */}
          <button
            onClick={handleCopy}
            disabled={selectedChapters.length === 0}
            style={{
              width: "100%", padding: "13px", borderRadius: 8, border: "none",
              background: selectedChapters.length > 0
                ? (copyStatus === "copied" ? "#059669" : "#1a56db")
                : "#e5e7eb",
              color: selectedChapters.length > 0 ? "#fff" : "#9ca3af",
              fontWeight: 700, fontSize: 15,
              cursor: selectedChapters.length > 0 ? "pointer" : "not-allowed",
              letterSpacing: 0.5,
              transition: "background 0.2s",
            }}
          >
            {copyStatus === "copied" ? "✅ 已複製！請貼至 Claude 對話框" :
             copyStatus === "show"   ? "📋 已全選！按 Ctrl+C（Mac: ⌘+C）複製" :
             "📋 複製配置到 Claude 對話框"}
          </button>

          {/* 配置文字框（顯示後自動全選，點擊再次全選） */}
          {(copyStatus === "show" || copyStatus === "copied") && selectedChapters.length > 0 && (
            <div style={{ marginTop: 10 }}>
              <div style={{
                fontSize: 12, color: "#1d4ed8", fontWeight: 600, marginBottom: 4,
                background: "#eff6ff", borderRadius: 4, padding: "4px 8px",
              }}>
                ✏️ 點擊下方文字框 → 全選（Ctrl+A）→ 複製（Ctrl+C）→ 貼回 Claude 對話框
              </div>
              <textarea
                ref={textareaRef}
                readOnly
                onClick={handleTextareaClick}
                value={buildCopyText()}
                rows={10}
                style={{
                  width: "100%", boxSizing: "border-box",
                  fontFamily: "monospace", fontSize: 12, lineHeight: 1.7,
                  padding: "10px 12px", borderRadius: 6,
                  border: "2px solid #3b82f6", background: "#f0f9ff",
                  color: "#1e3a5f", resize: "vertical", cursor: "text",
                }}
              />
            </div>
          )}
        </div>
      )}

      <div style={{ textAlign: "center", fontSize: 11, color: "#9ca3af", marginTop: 4 }}>
        品質計畫自動產生器 · Quality Plan Maker · Step 4
      </div>
    </div>
  );
}
