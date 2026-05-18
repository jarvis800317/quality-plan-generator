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
  amount: number,
  hasEquipment: boolean,
  checkedIds: Set<string>
): Chapter[] {
  return CHAPTERS.filter((ch) => {
    if (ch.id === 'ch6') {
      // 陸（設備功能運轉檢測程序及標準）：須同時滿足 hasEquipment + amount>=150
      return checkedIds.has('ch6') && hasEquipment && amount >= 150;
    }
    return checkedIds.has(ch.id);
  });
}