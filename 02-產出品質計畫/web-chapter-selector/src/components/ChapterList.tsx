import { CHAPTERS } from '../data/chapters';
import type { ChapterStatus } from '../data/chapters';
import { ChapterItem } from './ChapterItem';

interface ChapterListProps {
  chapterStates: Map<string, ChapterStatus>;
  checkedIds: Set<string>;       // 使用者已勾選的章節 ID
  onToggle: (id: string) => void;
}

export function ChapterList({ chapterStates, checkedIds, onToggle }: ChapterListProps) {
  return (
    <div className="flex flex-col gap-2">
      {CHAPTERS.map((ch) => {
        const status = chapterStates.get(ch.id) ?? 'optional';
        // hidden 狀態的章節不渲染
        if (status === 'hidden') return null;
        return (
          <ChapterItem
            key={ch.id}
            chapter={ch}
            status={status}
            isChecked={checkedIds.has(ch.id)}
            onToggle={onToggle}
          />
        );
      })}
    </div>
  );
}