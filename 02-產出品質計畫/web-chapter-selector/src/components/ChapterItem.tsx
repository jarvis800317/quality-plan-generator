import { Check } from 'lucide-react';
import clsx from 'clsx';
import type { Chapter, ChapterStatus } from '../data/chapters';

interface ChapterItemProps {
  chapter: Chapter;
  status: ChapterStatus;
  isChecked: boolean;          // 由外部（checkedIds）控制勾選狀態
  onToggle: (id: string) => void;
}

export function ChapterItem({ chapter, status, isChecked, onToggle }: ChapterItemProps) {
  const isRequired = status === 'required';
  // 必選章不可互動；選配章取決於 isChecked（由外部狀態控制）
  const isDisabled = isRequired;

  return (
    <div
      className={clsx(
        'chapter-row px-4 py-3 rounded-xl border text-sm transition-all',
        isRequired
          ? 'bg-slate-50 border-slate-200'
          : 'bg-white border-gray-100 hover:border-blue-200',
        isDisabled && 'opacity-80'
      )}
    >
      {/* 左側：勾選框 + 狀態標籤（水平排列，固定寬度） */}
      <div className="left-section">
        <button
          onClick={() => !isDisabled && onToggle(chapter.id)}
          disabled={isDisabled}
          className={clsx(
            'chk-box border-2 transition-all',
            isChecked && isRequired
              ? 'bg-blue-500 border-blue-500 text-white cursor-default'
              : isChecked
              ? 'bg-blue-500 border-blue-500 text-white cursor-pointer hover:bg-blue-600'
              : 'border-gray-300 bg-white cursor-pointer hover:border-blue-400',
            isDisabled ? 'cursor-default' : 'cursor-pointer'
          )}
          title={
            isRequired
              ? '必選章（依據契約金額門檻）'
              : '選配章（可自行勾選）'
          }
        >
          {isChecked && (
            <Check
              size={14}
              strokeWidth={3}
              color="black"
              style={{ display: 'block' }}
            />
          )}
        </button>
        <span
          className={clsx(
            'text-xs px-2 py-0.5 rounded-full font-semibold shrink-0',
            isRequired
              ? 'bg-blue-100 text-blue-600'
              : 'bg-amber-100 text-amber-600'
          )}
        >
          {isRequired ? '必選' : '選配'}
        </span>
      </div>

      {/* 中文大寫數字（固定寬度） */}
      <div className="numeral font-bold text-blue-600 text-base">
        {chapter.chineseNumeral.replace('、', '')}
      </div>

      {/* 章名（彈性延伸，不換行） */}
      <div className="title font-medium text-gray-800">
        {chapter.title}
      </div>
    </div>
  );
}