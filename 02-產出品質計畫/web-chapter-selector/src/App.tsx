import { useState, useCallback } from 'react';
import { ChapterList } from './components/ChapterList';
import { ProjectForm } from './components/ProjectForm';
import { OutputPanel } from './components/OutputPanel';
import { CHAPTERS, getChapterStates } from './data/chapters';
import { getSelectedChapters } from './hooks/useChapterFilter';
import type { ProjectCategory } from './hooks/useChapterFilter';
import appConfig from './config.json';
import './App.css';

// 預設值由 quality-plan-workflow.py 動態寫入 src/config.json
const DEFAULT_NAME = appConfig.projectName;
const DEFAULT_AMOUNT = appConfig.amountWan;

export default function App() {
  const [name, setName] = useState(DEFAULT_NAME);
  const [amount, setAmount] = useState(DEFAULT_AMOUNT);
  const [category, setCategory] = useState<ProjectCategory>('road');
  const [hasEquipment, setHasEquipment] = useState(false);

  // 各章節勾選狀態：必選章自動勾選，選配章預設不勾
  const [checkedIds, setCheckedIds] = useState<Set<string>>(() => {
    const amountNum = parseFloat(DEFAULT_AMOUNT) || 0;
    const states = getChapterStates(amountNum);
    const ids = new Set<string>();
    CHAPTERS.forEach((ch) => {
      if (states.get(ch.id) === 'required') ids.add(ch.id);
    });
    return ids;
  });

  // 依金額門檻重新計算必選章（保留使用者對選配章的手動操作）
  // 吃 amount 參數避免閉包陷阱；選配章保留原本勾選狀態
  const recompute = useCallback(
    (currentAmount: string) => {
      const amountNum = parseFloat(currentAmount) || 0;
      const states = getChapterStates(amountNum);
      setCheckedIds((prev) => {
        const next = new Set<string>();
        CHAPTERS.forEach((ch) => {
          const status = states.get(ch.id);
          if (status === 'required') {
            next.add(ch.id); // 必選章自動勾
          } else if (status === 'optional' && prev.has(ch.id)) {
            next.add(ch.id); // 選配章保留原本勾選狀態
          }
        });
        return next;
      });
    },
    []
  );

  function handleFormChange(field: string, value: string | boolean) {
    if (field === 'name') setName(value as string);
    if (field === 'amount') {
      setAmount(value as string);
      recompute(value as string); // 直接帶新金額，立即重新計算必選章
    }
    if (field === 'category') setCategory(value as ProjectCategory);
    if (field === 'hasEquipment') setHasEquipment(value as boolean);
  }

  // 切換選配章勾選狀態（hasEquipment 的限制僅影響輸出顯示，不阻止勾選）
  function handleToggle(id: string) {
    const amountNum = parseFloat(amount) || 0;
    const states = getChapterStates(amountNum);
    const status = states.get(id);
    // 只允許切換 optional 狀態的章節
    if (status !== 'optional') return;

    setCheckedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  const amountNum = parseFloat(amount) || 0;
  const chapterStates = getChapterStates(amountNum);
  const activeChapters = getSelectedChapters(amountNum, checkedIds);

  return (
    <div className="app-root">
      {/* 頁首 */}
      <header className="app-header">
        <div className="header-inner">
          <div>
            <h1 className="text-xl font-bold text-gray-800">
              品質計畫章節選擇器
            </h1>
            <p className="text-xs text-gray-400 mt-0.5">
              公共工程施工品質管理作業要點 · 品質計畫製作綱要（11章制）
            </p>
          </div>
          <div className="header-badge">MVP v0.2</div>
        </div>
      </header>

      <main className="app-main">
        {/* 左側：工程設定 + 章節列表 */}
        <div className="left-panel">
          <section className="card">
            <h2 className="card-title">工程基本資料</h2>
            <ProjectForm
              name={name}
              amount={amount}
              category={category}
              hasEquipment={hasEquipment}
              onChange={handleFormChange}
            />
            <button
              onClick={() => recompute(amount)}
              className="mt-4 w-full py-2 text-sm bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors"
            >
              重新計算章節
            </button>
          </section>

          <section className="card mt-4">
            <div className="flex items-center justify-between mb-3">
              <h2 className="card-title mb-0">章節列表（全11章）</h2>
              <span className="text-xs font-medium text-blue-600 bg-blue-50 px-2 py-1 rounded-full">
                {amountNum >= 5000
                  ? '11章制（10必選）'
                  : amountNum >= 1000
                  ? '11章制（6必選）'
                  : '11章制（4必選）'}
              </span>
            </div>
            <p className="text-xs text-gray-400 mb-3">
              必選章自動勾選，選配章可自行勾選
            </p>

            <ChapterList
              chapterStates={chapterStates}
              checkedIds={checkedIds}
              onToggle={handleToggle}
            />
          </section>
        </div>

        {/* 右側：輸出預覽 */}
        <div className="right-panel">
          <section className="card h-full">
            <h2 className="card-title">輸出預覽</h2>
            <div className="flex-1 min-h-0">
              <OutputPanel
                chapters={activeChapters}
                projectName={name || '未命名工程'}
                amount={amountNum}
              />
            </div>
          </section>
        </div>
      </main>
    </div>
  );
}