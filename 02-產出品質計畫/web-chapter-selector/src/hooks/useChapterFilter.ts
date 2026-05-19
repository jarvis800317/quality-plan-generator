import { CHAPTERS } from '../data/chapters';
import type { Chapter } from '../data/chapters';

export type ProjectCategory = 'building' | 'landscape' | 'road' | 'other';

export interface ProjectConfig {
  name: string;
  amount: number;    // 契約金額（萬）
  category: ProjectCategory;
  hasEquipment: boolean;
}

/**
 * 根據勾選狀態取得要輸出的章節列表（依中文大寫順序）
 */
export function getSelectedChapters(
  _amount: number,
  checkedIds: Set<string>
): Chapter[] {
  return CHAPTERS.filter((ch) => checkedIds.has(ch.id));
}