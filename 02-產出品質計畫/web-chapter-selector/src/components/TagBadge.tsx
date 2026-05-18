import clsx from 'clsx';
import type { ChapterStatus } from '../data/chapters';

interface TagBadgeProps {
  status: ChapterStatus;
  className?: string;
}

const LABELS: Record<ChapterStatus, string> = {
  required: '必選',
  optional: '選配',
  hidden: '隱藏',
};

const COLORS: Record<ChapterStatus, string> = {
  required: 'bg-blue-100 text-blue-700 border-blue-200',
  optional: 'bg-amber-50 text-amber-600 border-amber-200',
  hidden: 'bg-gray-100 text-gray-400 border-gray-200',
};

export function TagBadge({ status, className }: TagBadgeProps) {
  return (
    <span
      className={clsx(
        'inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold border',
        COLORS[status],
        className
      )}
    >
      {LABELS[status]}
    </span>
  );
}