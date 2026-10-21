/**
 * Type definitions for the Excel Akademi spreadsheet core
 */

export type HorizontalAlignment = 'left' | 'center' | 'right';
export type NumberFormat = 'general' | 'currency' | 'number' | 'percentage';

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  align?: HorizontalAlignment;
  format?: NumberFormat;
  textColor?: string;
  backgroundColor?: string;
  bgColor?: string;
}

export interface CellData {
  raw: string; // The user input or formula e.g. "=TOPLA(B2:B6)" or "Mouse" or "350"
  computed: string | number | null; // Evaluated display value
  style?: CellStyle;
  error?: string; // e.g. "#HATA!", "#DEĞER!"
}

export interface CellCoordinate {
  col: string;
  colIndex: number;
  row: number; // 1-indexed
}

export interface ColumnFilter {
  col?: string; // e.g. "A", "B"
  colLetter?: string;
  filterType?: 'text' | 'number' | 'values';
  operator?: 'equals' | 'contains' | 'greaterThan' | 'lessThan' | 'greaterOrEqual' | 'lessOrEqual' | 'notEquals' | string;
  condition?: string;
  value: string | number;
  selectedValues?: string[];
}

export interface SheetFilterConfig {
  headerRow?: number;
  startCol?: string;
  endCol?: string;
  startRow?: number;
  endRow?: number;
  filters: Record<string, ColumnFilter>;
  isActive: boolean;
}

export interface SheetChart {
  id: string;
  sheetId?: string;
  title: string;
  type?: 'column' | 'line' | 'pie';
  chartType?: 'column' | 'line' | 'pie';
  categoryRange: string; // e.g. "A2:A6"
  valueRange: string;    // e.g. "B2:B6"
  seriesName?: string;
  colorScheme?: string[];
  width?: number;
  height?: number;
}

export interface Worksheet {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key: e.g. "A1", "B7"
  rowCount: number;
  colCount: number;
  hiddenRows?: number[];
  filterConfig?: SheetFilterConfig;
  charts?: SheetChart[];
}

export interface Workbook {
  id: string;
  name: string;
  sheets: Worksheet[];
  activeSheetId: string;
}

export interface SelectionRange {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
}

export interface SelectionState {
  activeCell: CellCoordinate;
  anchorCell?: CellCoordinate;
  range: SelectionRange | null;
  isEditing: boolean;
  editValue: string;
}

export interface ClipboardData {
  isCut: boolean;
  sourceSheetId: string;
  sourceRange: SelectionRange;
  cells: Record<string, CellData>; // relative to sourceRange.startCol & startRow
  baseCell: CellCoordinate;
}

export interface ContextMenuState {
  isOpen: boolean;
  x: number;
  y: number;
  targetCell: CellCoordinate | null;
}

