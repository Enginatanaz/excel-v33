/**
 * Column Filter Dropdown Component
 * Provides Excel-style sorting and filtering controls directly on column headers
 */

import React, { useState } from 'react';
import { ArrowDownAZ, ArrowUpZA, Filter, X, Check } from 'lucide-react';
import type { ColumnFilter } from '../../modules/spreadsheet/types.ts';

interface ColumnFilterDropdownProps {
  colLetter: string;
  activeFilter?: ColumnFilter;
  onSort: (colLetter: string, ascending: boolean) => void;
  onApplyFilter: (colLetter: string, filter: ColumnFilter) => void;
  onClearFilter: (colLetter: string) => void;
  onClose: () => void;
}

export function ColumnFilterDropdown({
  colLetter,
  activeFilter,
  onSort,
  onApplyFilter,
  onClearFilter,
  onClose,
}: ColumnFilterDropdownProps) {
  const [filterInput, setFilterInput] = useState(activeFilter ? String(activeFilter.value) : '');

  const handleApply = (e: React.FormEvent) => {
    e.preventDefault();
    if (!filterInput.trim()) {
      onClearFilter(colLetter);
      onClose();
      return;
    }

    onApplyFilter(colLetter, {
      col: colLetter,
      filterType: 'text',
      value: filterInput.trim(),
    });
    onClose();
  };

  return (
    <div
      role="menu"
      className="absolute top-full left-0 z-50 mt-1 w-64 bg-white border border-slate-200 rounded-lg shadow-xl p-3 text-xs space-y-3 font-sans text-slate-800 animate-in fade-in duration-100"
      onClick={(e) => e.stopPropagation()}
    >
      <div className="flex items-center justify-between border-b border-slate-100 pb-2">
        <span className="font-bold text-slate-900 flex items-center gap-1.5">
          <Filter className="w-3.5 h-3.5 text-emerald-700" />
          <span>Sütun {colLetter} — Sırala & Filtrele</span>
        </span>
        <button
          onClick={onClose}
          className="p-0.5 text-slate-400 hover:text-slate-600 rounded cursor-pointer"
        >
          <X className="w-3.5 h-3.5" />
        </button>
      </div>

      {/* Sort Buttons */}
      <div className="space-y-1">
        <button
          onClick={() => {
            onSort(colLetter, true);
            onClose();
          }}
          className="w-full flex items-center gap-2 px-2.5 py-1.5 hover:bg-slate-100 rounded text-slate-700 text-left transition-colors cursor-pointer"
        >
          <ArrowDownAZ className="w-4 h-4 text-emerald-700" />
          <span>A'dan Z'ye Sırala (Artan)</span>
        </button>

        <button
          onClick={() => {
            onSort(colLetter, false);
            onClose();
          }}
          className="w-full flex items-center gap-2 px-2.5 py-1.5 hover:bg-slate-100 rounded text-slate-700 text-left transition-colors cursor-pointer"
        >
          <ArrowUpZA className="w-4 h-4 text-emerald-700" />
          <span>Z'den A'ya Sırala (Azalan)</span>
        </button>
      </div>

      <div className="border-t border-slate-100 pt-2 space-y-2">
        <label className="font-semibold text-slate-700 block">
          Koşula Göre Filtrele:
        </label>
        <form onSubmit={handleApply} className="space-y-2">
          <input
            type="text"
            value={filterInput}
            onChange={(e) => setFilterInput(e.target.value)}
            placeholder="Örn: >=1000 veya Ankara"
            className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-600 font-mono"
          />
          <div className="text-[10px] text-slate-400 leading-tight">
            Metin için <code>Ankara</code>, sayısal ölçüt için <code>&gt;=1000</code> yazabilirsiniz.
          </div>

          <div className="flex items-center justify-between pt-1">
            {activeFilter && (
              <button
                type="button"
                onClick={() => {
                  onClearFilter(colLetter);
                  onClose();
                }}
                className="text-[11px] text-rose-600 hover:text-rose-800 font-medium cursor-pointer"
              >
                Filtreyi Kaldır
              </button>
            )}
            <button
              type="submit"
              className="ml-auto flex items-center gap-1 px-3 py-1 bg-[#107c41] hover:bg-[#0b5a2f] text-white font-semibold rounded text-xs transition-colors cursor-pointer"
            >
              <Check className="w-3 h-3" />
              <span>Uygula</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
