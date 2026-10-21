/**
 * SVG-based Interactive Chart Viewer Component
 * Supports Column (Sütun), Line (Çizgi), and Pie (Pasta) charts
 * Real-time reactive binding to live worksheet cells
 */

import React, { useMemo } from 'react';
import type { SheetChart, Worksheet } from '../../modules/spreadsheet/types.ts';
import { parseRange, getAddressesInRange } from '../../utils/cellAddress.ts';
import { isNumeric, parseNumeric } from '../../utils/validation.ts';
import { BarChart3, LineChart, PieChart, Trash2, X } from 'lucide-react';

interface ChartViewerProps {
  chart: SheetChart;
  sheet: Worksheet;
  onDeleteChart?: (chartId: string) => void;
}

const DEFAULT_PALETTE = [
  '#107c41', // Excel green
  '#2563eb', // Blue
  '#f59e0b', // Amber
  '#ef4444', // Red
  '#8b5cf6', // Purple
  '#06b6d4', // Cyan
  '#ec4899', // Pink
  '#14b8a6', // Teal
];

export function ChartViewer({ chart, sheet, onDeleteChart }: ChartViewerProps) {
  // Extract live categories and values from worksheet cells
  const chartData = useMemo(() => {
    const catRange = parseRange(chart.categoryRange);
    const valRange = parseRange(chart.valueRange);

    const catAddrs = catRange ? getAddressesInRange(catRange) : [];
    const valAddrs = valRange ? getAddressesInRange(valRange) : [];

    const items: Array<{ category: string; value: number; formattedValue: string }> = [];

    const count = Math.min(catAddrs.length, valAddrs.length);
    for (let i = 0; i < count; i++) {
      const catCell = sheet.cells[catAddrs[i]];
      const valCell = sheet.cells[valAddrs[i]];

      const catStr = catCell?.computed !== null && catCell?.computed !== undefined
        ? String(catCell.computed)
        : catAddrs[i];

      const valRaw = valCell?.computed ?? valCell?.raw ?? 0;
      const valNum = isNumeric(valRaw) ? parseNumeric(valRaw) : 0;

      items.push({
        category: catStr,
        value: valNum,
        formattedValue: valNum.toLocaleString('tr-TR'),
      });
    }

    return items;
  }, [chart.categoryRange, chart.valueRange, sheet.cells]);

  const width = chart.width || 380;
  const height = chart.height || 230;
  const padding = { top: 35, right: 25, bottom: 45, left: 55 };

  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;

  const maxValue = useMemo(() => {
    const max = Math.max(...chartData.map((d) => d.value), 0);
    return max === 0 ? 100 : Math.ceil(max * 1.15);
  }, [chartData]);

  const totalValue = useMemo(() => {
    return chartData.reduce((acc, curr) => acc + (curr.value > 0 ? curr.value : 0), 0) || 1;
  }, [chartData]);

  return (
    <div className="relative bg-white border border-slate-200 rounded-xl p-3 shadow-sm hover:shadow-md transition-shadow">
      {/* Header & Controls */}
      <div className="flex items-center justify-between pb-2 border-b border-slate-100 mb-2">
        <div className="flex items-center gap-2">
          {chart.chartType === 'column' && <BarChart3 className="w-4 h-4 text-emerald-700" />}
          {chart.chartType === 'line' && <LineChart className="w-4 h-4 text-blue-600" />}
          {chart.chartType === 'pie' && <PieChart className="w-4 h-4 text-purple-600" />}
          <span className="font-bold text-xs sm:text-sm text-slate-800 truncate">
            {chart.title || 'Grafik'}
          </span>
          <span className="text-[10px] text-slate-400 font-mono">
            ({chart.categoryRange} & {chart.valueRange})
          </span>
        </div>

        {onDeleteChart && (
          <button
            onClick={() => onDeleteChart(chart.id)}
            title="Grafiği Sil"
            className="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition-colors cursor-pointer"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        )}
      </div>

      {chartData.length === 0 ? (
        <div className="h-44 flex items-center justify-center text-xs text-slate-400">
          Seçili aralıkta grafik verisi bulunamadı ({chart.categoryRange})
        </div>
      ) : (
        <div className="w-full overflow-x-auto flex justify-center">
          <svg
            viewBox={`0 0 ${width} ${height}`}
            className="w-full max-w-[420px] h-auto select-none font-sans"
            style={{ maxHeight: `${height}px` }}
          >
            {/* Background grid lines for Column and Line */}
            {chart.chartType !== 'pie' && (
              <g className="grid-lines" opacity={0.6}>
                {[0, 0.25, 0.5, 0.75, 1].map((pct, idx) => {
                  const y = padding.top + innerHeight * (1 - pct);
                  const tickVal = Math.round(maxValue * pct);
                  return (
                    <g key={idx}>
                      <line
                        x1={padding.left}
                        y1={y}
                        x2={width - padding.right}
                        y2={y}
                        stroke="#e2e8f0"
                        strokeDasharray={pct === 0 ? '' : '3 3'}
                        strokeWidth={pct === 0 ? 1.5 : 1}
                      />
                      <text
                        x={padding.left - 8}
                        y={y + 3.5}
                        textAnchor="end"
                        fontSize="10"
                        fill="#64748b"
                        className="font-mono"
                      >
                        {tickVal}
                      </text>
                    </g>
                  );
                })}
              </g>
            )}

            {/* A. SÜTUN GRAFİĞİ (Column Chart) */}
            {chart.chartType === 'column' && (
              <g className="columns">
                {chartData.map((d, idx) => {
                  const bandWidth = innerWidth / chartData.length;
                  const barWidth = Math.max(14, Math.min(bandWidth * 0.6, 40));
                  const x = padding.left + idx * bandWidth + (bandWidth - barWidth) / 2;
                  const barHeight = Math.max(2, (d.value / maxValue) * innerHeight);
                  const y = padding.top + innerHeight - barHeight;
                  const color = DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length];

                  return (
                    <g key={idx} className="group">
                      <rect
                        x={x}
                        y={y}
                        width={barWidth}
                        height={barHeight}
                        rx={3}
                        fill={color}
                        className="transition-all hover:opacity-85"
                      >
                        <title>{`${d.category}: ${d.formattedValue}`}</title>
                      </rect>
                      {/* Value label on top of bar */}
                      <text
                        x={x + barWidth / 2}
                        y={y - 5}
                        textAnchor="middle"
                        fontSize="9.5"
                        fontWeight="600"
                        fill="#334155"
                        className="font-mono"
                      >
                        {d.formattedValue}
                      </text>
                      {/* Category label below X axis */}
                      <text
                        x={x + barWidth / 2}
                        y={height - padding.bottom + 16}
                        textAnchor="middle"
                        fontSize="10"
                        fontWeight="500"
                        fill="#475569"
                        className="truncate max-w-[50px]"
                      >
                        {d.category.length > 8 ? d.category.slice(0, 7) + '…' : d.category}
                      </text>
                    </g>
                  );
                })}
              </g>
            )}

            {/* B. ÇİZGİ GRAFİĞİ (Line Chart) */}
            {chart.chartType === 'line' && (
              <g className="line-chart">
                {/* Area Gradient Defs */}
                <defs>
                  <linearGradient id={`grad_${chart.id}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#2563eb" stopOpacity="0.25" />
                    <stop offset="100%" stopColor="#2563eb" stopOpacity="0.0" />
                  </linearGradient>
                </defs>

                {(() => {
                  const points = chartData.map((d, idx) => {
                    const step = innerWidth / (chartData.length - 1 || 1);
                    const x = padding.left + idx * step;
                    const y = padding.top + innerHeight - (d.value / maxValue) * innerHeight;
                    return { x, y, data: d };
                  });

                  const pathD = points.reduce((acc, p, i) => `${acc} ${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`, '');
                  const areaD = `${pathD} L ${points[points.length - 1].x} ${padding.top + innerHeight} L ${points[0].x} ${padding.top + innerHeight} Z`;

                  return (
                    <>
                      <path d={areaD} fill={`url(#grad_${chart.id})`} />
                      <path d={pathD} fill="none" stroke="#2563eb" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />

                      {points.map((p, idx) => (
                        <g key={idx}>
                          <circle cx={p.x} cy={p.y} r={4.5} fill="#ffffff" stroke="#2563eb" strokeWidth="2.5" />
                          <text
                            x={p.x}
                            y={p.y - 8}
                            textAnchor="middle"
                            fontSize="9.5"
                            fontWeight="600"
                            fill="#1e293b"
                            className="font-mono"
                          >
                            {p.data.formattedValue}
                          </text>
                          <text
                            x={p.x}
                            y={height - padding.bottom + 16}
                            textAnchor="middle"
                            fontSize="10"
                            fontWeight="500"
                            fill="#475569"
                          >
                            {p.data.category.length > 8 ? p.data.category.slice(0, 7) + '…' : p.data.category}
                          </text>
                        </g>
                      ))}
                    </>
                  );
                })()}
              </g>
            )}

            {/* C. PASTA GRAFİĞİ (Pie Chart) */}
            {chart.chartType === 'pie' && (
              <g className="pie-chart">
                {(() => {
                  const cx = width * 0.38;
                  const cy = height * 0.52;
                  const radius = Math.min(innerWidth, innerHeight) * 0.44;

                  let accumulatedAngle = -Math.PI / 2; // start at top

                  const slices = chartData.map((d, idx) => {
                    const sliceAngle = (Math.max(d.value, 0) / totalValue) * (Math.PI * 2);
                    const startAngle = accumulatedAngle;
                    const endAngle = accumulatedAngle + sliceAngle;
                    accumulatedAngle = endAngle;

                    const x1 = cx + radius * Math.cos(startAngle);
                    const y1 = cy + radius * Math.sin(startAngle);
                    const x2 = cx + radius * Math.cos(endAngle);
                    const y2 = cy + radius * Math.sin(endAngle);

                    const largeArcFlag = sliceAngle > Math.PI ? 1 : 0;
                    const pathData = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;
                    const percent = Math.round((d.value / totalValue) * 100);
                    const color = DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length];

                    return { pathData, color, data: d, percent };
                  });

                  return (
                    <>
                      {slices.map((slice, idx) => (
                        <path
                          key={idx}
                          d={slice.pathData}
                          fill={slice.color}
                          stroke="#ffffff"
                          strokeWidth="2"
                          className="transition-all hover:opacity-90"
                        >
                          <title>{`${slice.data.category}: ${slice.data.formattedValue} (%${slice.percent})`}</title>
                        </path>
                      ))}

                      {/* Legend on right */}
                      <g transform={`translate(${width * 0.72}, 30)`}>
                        {chartData.map((d, idx) => {
                          const pct = Math.round((d.value / totalValue) * 100);
                          const color = DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length];
                          return (
                            <g key={idx} transform={`translate(0, ${idx * 22})`}>
                              <rect width="10" height="10" rx="2" fill={color} />
                              <text x="16" y="9" fontSize="10" fill="#334155" fontWeight="500">
                                {d.category} (%{pct})
                              </text>
                            </g>
                          );
                        })}
                      </g>
                    </>
                  );
                })()}
              </g>
            )}
          </svg>
        </div>
      )}
    </div>
  );
}
