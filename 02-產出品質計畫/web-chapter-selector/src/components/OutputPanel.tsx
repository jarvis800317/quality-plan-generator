import { useState } from 'react';
import { Download, Copy, Check } from 'lucide-react';
import type { Chapter } from '../data/chapters';
import { exportToText, downloadFile } from '../utils/exporters';

interface OutputPanelProps {
  chapters: Chapter[];
  projectName: string;
  amount: number;
}

export function OutputPanel({ chapters, projectName, amount }: OutputPanelProps) {
  const [copied, setCopied] = useState(false);

  const content = exportToText(chapters, projectName, amount);
  const filename = `${projectName.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_')}_品質計畫章節.txt`;

  function handleCopy() {
    navigator.clipboard.writeText(content).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  }

  function handleDownload() {
    downloadFile(content, filename);
  }

  return (
    <div className="flex flex-col h-full gap-3">
      <pre className="output-pre p-4 text-sm leading-relaxed text-gray-700 bg-gray-50 border border-gray-200 rounded-xl whitespace-pre-wrap font-sans">
        {content}
      </pre>

      <div className="output-actions">
        <button
          onClick={handleCopy}
          className="flex items-center gap-1.5 px-4 py-2 text-sm bg-white border border-gray-200 rounded-lg text-gray-600 hover:bg-gray-50 hover:border-gray-300 transition-colors"
        >
          {copied ? (
            <Check size={14} className="text-green-500" />
          ) : (
            <Copy size={14} />
          )}
          {copied ? '已複製' : '複製'}
        </button>
        <button
          onClick={handleDownload}
          disabled={chapters.length === 0}
          className="flex items-center gap-1.5 px-4 py-2 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed font-medium"
        >
          <Download size={14} />
          下載章節 (.txt)
        </button>
        <button
          onClick={() => {
            window.dispatchEvent(new CustomEvent('quality-plan-trigger', {
              detail: {
                projectName,
                amount,
                selectedIds: chapters.map((ch) => ch.id),
                timestamp: Date.now(),
              },
            }));
          }}
          disabled={chapters.length === 0}
          className="flex items-center gap-1.5 px-4 py-2 text-sm bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed font-medium"
        >
          ✓ 確認並產生品質計畫
        </button>
        <span className="flex-1" />
        <span className="flex items-center text-xs text-gray-400 self-center">
          共 {chapters.length} 章
        </span>
      </div>
    </div>
  );
}