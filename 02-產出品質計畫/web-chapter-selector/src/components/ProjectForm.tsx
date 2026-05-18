import type { ProjectCategory } from '../hooks/useChapterFilter';

interface ProjectFormProps {
  name: string;
  amount: string;
  category: ProjectCategory;
  hasEquipment: boolean;
  onChange: (field: string, value: string | boolean) => void;
}

const CATEGORIES: { value: ProjectCategory; label: string }[] = [
  { value: 'building', label: '建築' },
  { value: 'landscape', label: '景觀' },
  { value: 'road', label: '道路' },
  { value: 'other', label: '其他' },
];

export function ProjectForm({
  name,
  amount,
  category,
  hasEquipment,
  onChange,
}: ProjectFormProps) {
  return (
    <div className="flex flex-col gap-4">
      {/* 工程名稱 */}
      <div>
        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1.5">
          工程名稱
        </label>
        <input
          type="text"
          value={name}
          onChange={(e) => onChange('name', e.target.value)}
          placeholder="例如：桃林鐵路鋪面改善工程"
          className="w-full px-3 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-400 focus:border-transparent"
        />
      </div>

      {/* 契約金額 */}
      <div>
        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1.5">
          契約金額（萬元）
        </label>
        <input
          type="number"
          value={amount}
          onChange={(e) => onChange('amount', e.target.value)}
          placeholder="例如：4600"
          min="0"
          className="w-full px-3 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-400 focus:border-transparent font-mono"
        />
      </div>

      {/* 工程類別 */}
      <div>
        <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1.5">
          工程類別
        </label>
        <div className="flex flex-wrap gap-2">
          {CATEGORIES.map((cat) => (
            <button
              key={cat.value}
              onClick={() => onChange('category', cat.value)}
              className={`px-3 py-1.5 text-sm rounded-lg border transition-all ${
                category === cat.value
                  ? 'bg-blue-50 border-blue-400 text-blue-700 font-medium'
                  : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
              }`}
            >
              {cat.label}
            </button>
          ))}
        </div>
      </div>

      {/* 機電設備 */}
      <div className="flex items-center gap-3 p-3 bg-amber-50 border border-amber-200 rounded-lg">
        <input
          type="checkbox"
          id="hasEquipment"
          checked={hasEquipment}
          onChange={(e) => onChange('hasEquipment', e.target.checked)}
          className="w-4 h-4 accent-amber-600"
        />
        <label htmlFor="hasEquipment" className="text-sm text-amber-800 cursor-pointer">
          本工程有<strong>運轉類機電設備</strong>（照明監控、排水設備、噴灌系統等）
        </label>
      </div>
    </div>
  );
}