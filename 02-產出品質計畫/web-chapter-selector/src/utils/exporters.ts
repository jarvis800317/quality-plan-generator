import type { Chapter } from '../data/chapters';

// 中文數字對照表
const CHINESE_NUMS = ['壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖', '拾', '拾壹', '拾貳', '拾參'];

// 產生輸出文件（選中章節依序重新編號：壹、貳、參…而非原始數字）
export function exportToText(
  chapters: Chapter[],
  projectName: string,
  amount: number
): string {
  if (chapters.length === 0) {
    return `工程名稱：${projectName}\n契約金額：${amount}萬元\n（未勾選任何章節）`;
  }

  const lines = [
    `工程名稱：${projectName}`,
    `契約金額：${amount}萬元`,
    '',
    ...chapters.map((ch, i) => `${CHINESE_NUMS[i]}、${ch.title}`),
  ];

  return lines.join('\n');
}

// 下載工具
export function downloadFile(content: string, filename: string) {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}