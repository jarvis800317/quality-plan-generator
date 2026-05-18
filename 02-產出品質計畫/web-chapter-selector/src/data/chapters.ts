// 品質計畫製作綱要 — 11章結構（正確順序）
// 章名與順序依「品質計畫製作綱要」規定

export type ChapterStatus = 'required' | 'optional' | 'hidden';

export interface Chapter {
  id: string;
  chineseNumeral: string; // 壹、貳、參...
  title: string;          // 綱要規定章名
}

// 正確順序（共11章，設備功能運轉檢測程序及標準提前至第6位）
export const CHAPTERS: Chapter[] = [
  { id: 'ch1',  chineseNumeral: '壹',  title: '計畫範圍' },
  { id: 'ch2',  chineseNumeral: '貳',  title: '管理權責及分工' },
  { id: 'ch3',  chineseNumeral: '參',  title: '施工要領' },
  { id: 'ch4',  chineseNumeral: '肆',  title: '品質管理標準' },
  { id: 'ch5',  chineseNumeral: '伍',  title: '材料與設備及施工檢驗程序' },
  { id: 'ch6',  chineseNumeral: '陸',  title: '設備功能運轉檢測程序及標準' },
  { id: 'ch7',  chineseNumeral: '柒',  title: '自主檢查表' },
  { id: 'ch8',  chineseNumeral: '捌',  title: '不合格品之管制' },
  { id: 'ch9',  chineseNumeral: '玖',  title: '矯正與預防措施' },
  { id: 'ch10', chineseNumeral: '拾',  title: '內部品質稽核' },
  { id: 'ch11', chineseNumeral: '拾壹', title: '文件紀錄管理系統' },
];

// 三個級距門檻（萬）
export const THRESHOLDS = {
  TIER_A: 5000, // 5,000萬以上：10必選
  TIER_B: 1000, // 1,000萬～<5,000萬：6必選
  TIER_C: 150,  // 150萬～<1,000萬：4必選
};

// 各章在三個級距下的狀態
// TIER_A(col0, ≥5000萬): 10必選 = 壹至伍全部(ch1-ch5) + 柒至拾壹全部(ch7-ch11)，僅陸(ch6)選配
// TIER_B(col1, ≥1000萬且<5000萬): 6必選（用戶確認：壹、貳、肆、伍、柒、拾壹）
// TIER_C(col2, ≥150萬且<1000萬): 4必選（用戶列：貳、肆、伍、柒）
const STATUS_TABLE: [string, ChapterStatus, ChapterStatus, ChapterStatus][] = [
  ['ch1',  'required', 'required',  'optional'],   // 壹：計畫範圍（TIER_B確認必選）
  ['ch2',  'required', 'required',  'required'],   // 貳：管理權責及分工
  ['ch3',  'required', 'optional',  'optional'],   // 參：施工要領
  ['ch4',  'required', 'required',  'required'],   // 肆：品質管理標準
  ['ch5',  'required', 'required',  'required'],   // 伍：材料與設備及施工檢驗程序
  ['ch6',  'optional', 'optional',  'optional'],   // 陸：設備功能運轉檢測程序及標準（永遠選配）
  ['ch7',  'required', 'required',  'required'],   // 柒：自主檢查表
  ['ch8',  'required', 'optional',  'optional'],   // 捌：不合格品之管制
  ['ch9',  'required', 'optional',  'optional'],   // 玖：矯正與預防措施
  ['ch10', 'required', 'optional',  'optional'],   // 拾：內部品質稽核
  ['ch11', 'required', 'required',  'optional'],   // 拾壹：文件紀錄管理系統
];

/**
 * 根據契約金額決定使用哪一個級距的狀態對照表
 * col 0 = TIER_A (≥5000萬, 10必選), col 1 = TIER_B (≥1000萬, 6必選), col 2 = TIER_C (<1000萬, 4必選)
 */
export function getStatusMap(amount: number): Map<string, ChapterStatus> {
  const col = amount >= THRESHOLDS.TIER_A ? 0
             : amount >= THRESHOLDS.TIER_B ? 1 : 2;

  const result = new Map<string, ChapterStatus>();
  for (const [id, s0, s1, s2] of STATUS_TABLE) {
    result.set(id, [s0, s1, s2][col] as ChapterStatus);
  }
  return result;
}

/**
 * 計算每一章的狀態
 * - required：必選（無方框，不可取消）
 * - optional：選配（有空方框，可自由勾選）
 */
export function getChapterStates(amount: number): Map<string, ChapterStatus> {
  return getStatusMap(amount);
}

/**
 * 計算應顯示的級距標記文字
 */
export function getTierLabel(amount: number): string {
  if (amount >= THRESHOLDS.TIER_A) return '11章制（5,000萬以上）';
  if (amount >= THRESHOLDS.TIER_B) return '11章制（1,000萬～<5,000萬）';
  return '11章制（150萬～<1,000萬）';
}