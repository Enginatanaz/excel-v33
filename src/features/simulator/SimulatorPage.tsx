/**
 * Standalone Excel Simulator Workbench
 * Full-width interactive spreadsheet sandbox with formulas, sorting, filtering, charting, toolbar, and sheet tabs
 */

import { useState } from 'react';
import { useSpreadsheet } from '../../hooks/useSpreadsheet.ts';
import { ExcelToolbar } from '../../components/excel/ExcelToolbar.tsx';
import { FormulaBar } from '../../components/excel/FormulaBar.tsx';
import { SpreadsheetGrid } from '../../components/excel/SpreadsheetGrid.tsx';
import { SheetTabs } from '../../components/excel/SheetTabs.tsx';
import { CellContextMenu } from '../../components/excel/CellContextMenu.tsx';
import { ChartViewer } from '../../components/excel/ChartViewer.tsx';
import { CreateChartModal } from '../../components/excel/CreateChartModal.tsx';
import { createDefaultWorkbook } from '../../modules/spreadsheet/spreadsheetState.ts';

export function SimulatorPage() {
  const spreadsheet = useSpreadsheet(createDefaultWorkbook);
  const [isChartModalOpen, setIsChartModalOpen] = useState(false);

  const activeCharts = spreadsheet.activeSheet.charts || [];

  return (
    <div className="space-y-4">
      {/* Page Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2 border-b border-slate-200">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight">
            Excel Çalışma Alanı Simülatörü
          </h1>
          <p className="text-xs sm:text-sm text-slate-500">
            Aşama 5: EĞER, DÜŞEYARA, ÇAPRAZARA formülleri, A-Z sıralama, filtreleme ve dinamik grafik motoru.
          </p>
        </div>

        <div className="text-xs font-mono text-slate-500 bg-white px-3 py-1.5 rounded border border-slate-200 self-start sm:self-auto">
          Aktif: {spreadsheet.activeCellAddress}
        </div>
      </div>

      {/* Spreadsheet Workspace Container */}
      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden shadow-xs">
        {/* Ribbon / Toolbar */}
        <ExcelToolbar
          activeCellStyle={spreadsheet.activeCellData?.style}
          onToggleBold={() =>
            spreadsheet.toggleStyle({
              bold: !spreadsheet.activeCellData?.style?.bold,
            })
          }
          onToggleItalic={() =>
            spreadsheet.toggleStyle({
              italic: !spreadsheet.activeCellData?.style?.italic,
            })
          }
          onSetAlignment={(align) => spreadsheet.toggleStyle({ align })}
          onToggleCurrency={() =>
            spreadsheet.toggleStyle({
              format:
                spreadsheet.activeCellData?.style?.format === 'currency'
                  ? 'general'
                  : 'currency',
            })
          }
          onAutoSum={() => spreadsheet.applyAutoSum()}
          onCopy={spreadsheet.copyCell}
          onCut={spreadsheet.cutCell}
          onPaste={spreadsheet.pasteCell}
          onClear={spreadsheet.clearCell}
          onReset={() => spreadsheet.resetWorkbook()}
          onUndo={spreadsheet.undo}
          onRedo={spreadsheet.redo}
          canUndo={spreadsheet.canUndo}
          canRedo={spreadsheet.canRedo}
          hasClipboardData={!!spreadsheet.clipboard}
          onSortAscending={() => spreadsheet.sortColumn(spreadsheet.selection.activeCell.col, true)}
          onSortDescending={() => spreadsheet.sortColumn(spreadsheet.selection.activeCell.col, false)}
          onToggleFilter={spreadsheet.toggleFilterMode}
          isFilterActive={Boolean(spreadsheet.activeSheet.filterConfig?.isActive)}
          onOpenCreateChart={() => setIsChartModalOpen(true)}
        />

        {/* Formula Bar */}
        <FormulaBar
          activeCellAddress={spreadsheet.activeCellAddress}
          displayAddress={
            spreadsheet.selection.range
              ? `${spreadsheet.selection.range.startCol}${spreadsheet.selection.range.startRow}:${spreadsheet.selection.range.endCol}${spreadsheet.selection.range.endRow}`
              : undefined
          }
          rawValue={spreadsheet.activeCellData?.raw || ''}
          isEditing={spreadsheet.selection.isEditing}
          editValue={spreadsheet.selection.editValue}
          onStartEditing={spreadsheet.startEditing}
          onUpdateEditValue={spreadsheet.updateEditValue}
          onCommitEditing={spreadsheet.commitEditing}
          onCancelEditing={spreadsheet.cancelEditing}
          onJumpToAddress={spreadsheet.jumpToAddress}
        />

        {/* Grid */}
        <SpreadsheetGrid
          sheet={spreadsheet.activeSheet}
          selection={spreadsheet.selection}
          onSelectCell={spreadsheet.selectCell}
          onSelectAll={spreadsheet.selectAll}
          onStartEditing={spreadsheet.startEditing}
          onUpdateEditValue={spreadsheet.updateEditValue}
          onCommitEditing={spreadsheet.commitEditing}
          onCancelEditing={spreadsheet.cancelEditing}
          onClearCell={spreadsheet.clearCell}
          onContextMenu={(e, cell) => spreadsheet.openContextMenu(e.clientX, e.clientY, cell)}
          onCopy={spreadsheet.copyCell}
          onCut={spreadsheet.cutCell}
          onPaste={spreadsheet.pasteCell}
          onUndo={spreadsheet.undo}
          onRedo={spreadsheet.redo}
          onToggleBold={() =>
            spreadsheet.toggleStyle({
              bold: !spreadsheet.activeCellData?.style?.bold,
            })
          }
          onSave={spreadsheet.saveWorkbook}
          onSortColumn={spreadsheet.sortColumn}
          onApplyFilter={spreadsheet.applyFilter}
          onClearFilter={spreadsheet.clearFilter}
        />

        {/* Sheet Tabs */}
        <SheetTabs
          sheets={spreadsheet.workbook.sheets}
          activeSheetId={spreadsheet.activeSheetId}
          onSwitchSheet={spreadsheet.switchSheet}
          onAddSheet={spreadsheet.addSheet}
          onRenameSheet={spreadsheet.renameSheet}
          onDeleteSheet={spreadsheet.deleteSheet}
        />
      </div>

      {/* Render Active Charts for this sheet */}
      {activeCharts.length > 0 && (
        <div className="space-y-3 pt-2">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-slate-800">
              Çalışma Sayfası Grafikleri ({activeCharts.length})
            </h3>
            <button
              onClick={() => setIsChartModalOpen(true)}
              className="text-xs text-emerald-700 hover:text-emerald-800 font-medium cursor-pointer"
            >
              + Yeni Grafik Ekle
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {activeCharts.map((chart) => (
              <ChartViewer
                key={chart.id}
                chart={chart}
                sheet={spreadsheet.activeSheet}
                onDeleteChart={spreadsheet.deleteChart}
              />
            ))}
          </div>
        </div>
      )}

      {/* Create Chart Modal */}
      <CreateChartModal
        isOpen={isChartModalOpen}
        onClose={() => setIsChartModalOpen(false)}
        onCreateChart={spreadsheet.addChart}
        selectedRange={spreadsheet.selection.range}
      />

      {/* Floating Status Notification */}
      {spreadsheet.statusMessage && (
        <div className="fixed bottom-14 right-4 sm:bottom-6 sm:right-6 bg-slate-900 text-white px-3 py-1.5 rounded-md shadow-lg text-xs font-medium z-50 animate-in fade-in slide-in-from-bottom-2">
          {spreadsheet.statusMessage}
        </div>
      )}

      {/* Quick Help Footer */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs text-slate-500 pt-2">
        <div className="p-3 bg-white border border-slate-200 rounded-md">
          <span className="font-semibold text-slate-800 block mb-1">Gelişmiş Formüller</span>
          <code className="text-emerald-700 font-mono">=EĞER(B2&gt;=50;"Geçti";"Kaldı")</code>, <code className="text-emerald-700 font-mono">=EĞERSAY(A2:A6;"Ankara")</code>, <code className="text-emerald-700 font-mono">=DÜŞEYARA(A2;F2:H6;2;0)</code>, <code className="text-emerald-700 font-mono">=ÇAPRAZARA(...)</code>
        </div>
        <div className="p-3 bg-white border border-slate-200 rounded-md">
          <span className="font-semibold text-slate-800 block mb-1">Sıralama & Filtreleme</span>
          Araç çubuğundaki A-Z / Z-A butonları ile satır bütünlüğünü koruyarak sıralayabilir, Filtre butonuyla sütun başlıklarında açılır filtreleri açabilirsiniz.
        </div>
        <div className="p-3 bg-white border border-slate-200 rounded-md">
          <span className="font-semibold text-slate-800 block mb-1">Mini Grafik Motoru</span>
          "Grafik Ekle" butonuna basarak seçili aralıktan canlı sütun, çizgi veya pasta grafiği üretebilirsiniz. Hücreler değiştikçe grafikler anında güncellenir.
        </div>
      </div>

      {/* Context Menu */}
      <CellContextMenu
        contextMenu={spreadsheet.contextMenu}
        onClose={spreadsheet.closeContextMenu}
        onEdit={() => spreadsheet.startEditing()}
        onCopy={spreadsheet.copyCell}
        onCut={spreadsheet.cutCell}
        onPaste={spreadsheet.pasteCell}
        onClear={spreadsheet.clearCell}
        hasClipboard={!!spreadsheet.clipboard}
      />
    </div>
  );
}
