/**
 * Create Chart Modal Component
 * Dialog to configure and insert Column, Line, or Pie charts
 */

import React, { useState, useEffect } from 'react';
import { BarChart3, LineChart, PieChart, X, Check } from 'lucide-react';
import type { SheetChart, SelectionRange } from '../../modules/spreadsheet/types.ts';

interface CreateChartModalProps {
  isOpen: boolean;
  onClose: () => void;
  onCreateChart: (chart: Omit<SheetChart, 'id' | 'sheetId'>) => void;
  selectedRange: SelectionRange | null;
}

export function CreateChartModal({
  isOpen,
  onClose,
  onCreateChart,
  selectedRange,
}: CreateChartModalProps) {
  const [chartType, setChartType] = useState<'column' | 'line' | 'pie'>('column');
  const [title, setTitle] = useState('Veri Grafiği');
  const [categoryRange, setCategoryRange] = useState('A2:A6');
  const [valueRange, setValueRange] = useState('B2:B6');

  // Auto-populate from selection if user had a range selected
  useEffect(() => {
    if (selectedRange) {
      if (selectedRange.startCol !== selectedRange.endCol) {
        // Multi-column selection (e.g. A2:B6)
        setCategoryRange(`${selectedRange.startCol}${selectedRange.startRow}:${selectedRange.startCol}${selectedRange.endRow}`);
        setValueRange(`${selectedRange.endCol}${selectedRange.startRow}:${selectedRange.endCol}${selectedRange.endRow}`);
      } else {
        // Single column selected -> set as values
        setValueRange(`${selectedRange.startCol}${selectedRange.startRow}:${selectedRange.endCol}${selectedRange.endRow}`);
      }
    }
  }, [selectedRange]);

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!categoryRange || !valueRange) return;

    onCreateChart({
      title: title.trim() || 'Veri Grafiği',
      chartType,
      categoryRange: categoryRange.trim().toUpperCase(),
      valueRange: valueRange.trim().toUpperCase(),
    });

    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-xl shadow-xl max-w-md w-full p-5 space-y-4 text-slate-800"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-slate-200 pb-3">
          <div className="flex items-center gap-2">
            <span className="p-1.5 bg-emerald-100 text-emerald-800 rounded-lg">
              <BarChart3 className="w-5 h-5" />
            </span>
            <h3 className="font-bold text-base text-slate-900">Grafik Oluştur</h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-slate-400 hover:text-slate-600 rounded-md cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4 text-xs">
          {/* Chart Type Picker */}
          <div>
            <label className="font-semibold text-slate-700 block mb-2">Grafik Türü Seçin:</label>
            <div className="grid grid-cols-3 gap-2">
              <button
                type="button"
                onClick={() => setChartType('column')}
                className={`flex flex-col items-center justify-center gap-1.5 p-3 rounded-lg border text-center transition-all cursor-pointer ${
                  chartType === 'column'
                    ? 'border-emerald-600 bg-emerald-50 text-emerald-900 font-bold ring-2 ring-emerald-500/20'
                    : 'border-slate-200 hover:border-slate-300 text-slate-700'
                }`}
              >
                <BarChart3 className="w-6 h-6 text-emerald-700" />
                <span>Sütun Grafiği</span>
              </button>

              <button
                type="button"
                onClick={() => setChartType('line')}
                className={`flex flex-col items-center justify-center gap-1.5 p-3 rounded-lg border text-center transition-all cursor-pointer ${
                  chartType === 'line'
                    ? 'border-blue-600 bg-blue-50 text-blue-900 font-bold ring-2 ring-blue-500/20'
                    : 'border-slate-200 hover:border-slate-300 text-slate-700'
                }`}
              >
                <LineChart className="w-6 h-6 text-blue-600" />
                <span>Çizgi Grafiği</span>
              </button>

              <button
                type="button"
                onClick={() => setChartType('pie')}
                className={`flex flex-col items-center justify-center gap-1.5 p-3 rounded-lg border text-center transition-all cursor-pointer ${
                  chartType === 'pie'
                    ? 'border-purple-600 bg-purple-50 text-purple-900 font-bold ring-2 ring-purple-500/20'
                    : 'border-slate-200 hover:border-slate-300 text-slate-700'
                }`}
              >
                <PieChart className="w-6 h-6 text-purple-600" />
                <span>Pasta Grafiği</span>
              </button>
            </div>
          </div>

          {/* Title */}
          <div>
            <label className="font-semibold text-slate-700 block mb-1">Grafik Başlığı:</label>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="Örn: 2025 Ürün Satış Dağılımı"
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-600 font-medium"
            />
          </div>

          {/* Category Range */}
          <div>
            <label className="font-semibold text-slate-700 block mb-1">
              Kategori / Etiket Aralığı (X Ekseni):
            </label>
            <input
              type="text"
              value={categoryRange}
              onChange={(e) => setCategoryRange(e.target.value)}
              placeholder="Örn: A2:A6"
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-600 font-mono font-medium"
            />
            <span className="text-[11px] text-slate-400 mt-0.5 block">
              Ürün adları, şehirler veya aylar gibi metin etiketleri
            </span>
          </div>

          {/* Value Range */}
          <div>
            <label className="font-semibold text-slate-700 block mb-1">
              Sayısal Değer Aralığı (Y Ekseni):
            </label>
            <input
              type="text"
              value={valueRange}
              onChange={(e) => setValueRange(e.target.value)}
              placeholder="Örn: B2:B6"
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-600 font-mono font-medium"
            />
            <span className="text-[11px] text-slate-400 mt-0.5 block">
              Grafikte çizilecek sayısal tutarlar
            </span>
          </div>

          <div className="flex justify-end gap-2 pt-2 border-t border-slate-200">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-1.5 border border-slate-300 hover:bg-slate-50 rounded-lg text-slate-700 font-medium cursor-pointer"
            >
              İptal
            </button>
            <button
              type="submit"
              className="flex items-center gap-1.5 px-4 py-1.5 bg-[#107c41] hover:bg-[#0b5a2f] text-white font-semibold rounded-lg shadow-xs cursor-pointer"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Grafiği Ekle</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
