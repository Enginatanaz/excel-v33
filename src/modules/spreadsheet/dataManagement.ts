/**
 * Data Management Engine for Excel Akademi AI
 * Stage 5: Sorting, Multi-Column Filtering, Range Detection and Row Integrity
 */

import type { Worksheet, SheetFilterConfig, ColumnFilter, CellData } from './types.ts';
import {
  letterToColIndex,
  colIndexToLetter,
  cellToAddress,
  addressToCell,
} from '../../utils/cellAddress.ts';
import { isNumeric, parseNumeric } from '../../utils/validation.ts';
import { matchesCriterion } from './formulaEngine.ts';

export interface TableBounds {
  startColIndex: number;
  endColIndex: number;
  startRow: number;
  endRow: number;
  headerRow: number;
}

/**
 * Automatically detects the contiguous rectangular table region around an active cell or the sheet
 */
export function detectTableBounds(
  sheet: Worksheet,
  seedCell?: { colIndex: number; row: number }
): TableBounds | null {
  const cellKeys = Object.keys(sheet.cells).filter((k) => {
    const c = sheet.cells[k];
    return c && (c.raw !== '' || c.computed !== null);
  });

  if (cellKeys.length === 0) return null;

  let minCol = Infinity;
  let maxCol = -Infinity;
  let minRow = Infinity;
  let maxRow = -Infinity;

  for (const k of cellKeys) {
    const parsed = addressToCell(k);
    if (!parsed) continue;
    if (parsed.colIndex < minCol) minCol = parsed.colIndex;
    if (parsed.colIndex > maxCol) maxCol = parsed.colIndex;
    if (parsed.row < minRow) minRow = parsed.row;
    if (parsed.row > maxRow) maxRow = parsed.row;
  }

  if (minCol === Infinity || minRow === Infinity) return null;

  return {
    startColIndex: minCol,
    endColIndex: maxCol,
    headerRow: minRow,
    startRow: minRow + 1, // Data starts below header
    endRow: maxRow,
  };
}

/**
 * Sorts worksheet data by a specific column while strictly preserving ROW INTEGRITY
 * When a row is sorted, all corresponding cells in that row move together!
 * Header row is kept fixed.
 */
export function sortWorksheet(
  sheet: Worksheet,
  targetCol: string,
  ascending: boolean = true,
  explicitBounds?: TableBounds
): Worksheet {
  const bounds = explicitBounds || detectTableBounds(sheet);
  if (!bounds || bounds.startRow > bounds.endRow) {
    return sheet;
  }

  const targetColIndex = letterToColIndex(targetCol);
  if (targetColIndex < bounds.startColIndex || targetColIndex > bounds.endColIndex) {
    return sheet;
  }

  // 1. Collect all rows from startRow to endRow
  interface RowRecord {
    originalRow: number;
    sortKeyVal: any;
    cells: Record<string, CellData>; // key is column letter
  }

  const rows: RowRecord[] = [];

  for (let r = bounds.startRow; r <= bounds.endRow; r++) {
    const sortAddr = cellToAddress(targetColIndex, r);
    const sortCell = sheet.cells[sortAddr];
    const rawVal = sortCell?.computed ?? sortCell?.raw ?? '';

    const rowCells: Record<string, CellData> = {};
    for (let c = bounds.startColIndex; c <= bounds.endColIndex; c++) {
      const colLetter = colIndexToLetter(c);
      const addr = cellToAddress(c, r);
      if (sheet.cells[addr]) {
        rowCells[colLetter] = { ...sheet.cells[addr] };
      }
    }

    rows.push({
      originalRow: r,
      sortKeyVal: rawVal,
      cells: rowCells,
    });
  }

  // 2. Sort row records
  rows.sort((a, b) => {
    const valA = a.sortKeyVal;
    const valB = b.sortKeyVal;

    // Empty cells always sink to bottom regardless of order
    const isEmptyA = valA === null || valA === undefined || valA === '';
    const isEmptyB = valB === null || valB === undefined || valB === '';
    if (isEmptyA && isEmptyB) return 0;
    if (isEmptyA) return 1;
    if (isEmptyB) return -1;

    const isNumA = isNumeric(valA);
    const isNumB = isNumeric(valB);

    if (isNumA && isNumB) {
      const diff = parseNumeric(valA) - parseNumeric(valB);
      return ascending ? diff : -diff;
    }

    const strA = String(valA).toLocaleLowerCase('tr-TR');
    const strB = String(valB).toLocaleLowerCase('tr-TR');
    const cmp = strA.localeCompare(strB, 'tr-TR');
    return ascending ? cmp : -cmp;
  });

  // 3. Reconstruct worksheet cells
  const newCells = { ...sheet.cells };

  // Clear existing row space
  for (let r = bounds.startRow; r <= bounds.endRow; r++) {
    for (let c = bounds.startColIndex; c <= bounds.endColIndex; c++) {
      const addr = cellToAddress(c, r);
      delete newCells[addr];
    }
  }

  // Write sorted rows back into consecutive rows
  rows.forEach((rowRecord, idx) => {
    const newRow = bounds.startRow + idx;
    for (let c = bounds.startColIndex; c <= bounds.endColIndex; c++) {
      const colLetter = colIndexToLetter(c);
      const newAddr = cellToAddress(c, newRow);
      if (rowRecord.cells[colLetter]) {
        newCells[newAddr] = rowRecord.cells[colLetter];
      }
    }
  });

  return {
    ...sheet,
    cells: newCells,
  };
}

/**
 * Applies multi-column filtering to the worksheet
 * Calculates hiddenRows (rows that do not match the criteria)
 */
export function applySheetFilters(
  sheet: Worksheet,
  config: SheetFilterConfig
): Worksheet {
  if (!config.isActive || Object.keys(config.filters).length === 0) {
    return {
      ...sheet,
      hiddenRows: [],
      filterConfig: { ...config, isActive: false },
    };
  }

  const bounds = detectTableBounds(sheet);
  const startRow = config.startRow ?? (bounds ? bounds.startRow : 2);
  const endRow = config.endRow ?? (bounds ? bounds.endRow : sheet.rowCount);

  const hidden: number[] = [];

  for (let r = startRow; r <= endRow; r++) {
    let rowMatches = true;

    for (const [colLetter, filter] of Object.entries(config.filters)) {
      if (filter.condition === 'none') continue;

      const addr = `${colLetter}${r}`;
      const cellVal = sheet.cells[addr]?.computed ?? sheet.cells[addr]?.raw ?? '';

      // Checklist filtering (selectedValues)
      if (filter.selectedValues && filter.selectedValues.length > 0) {
        const strVal = String(cellVal ?? '');
        if (!filter.selectedValues.includes(strVal)) {
          rowMatches = false;
          break;
        }
      }

      // Criterion filtering
      if (filter.value !== undefined && filter.value !== '') {
        const valStr = String(filter.value).trim();
        const hasOp = /^[><=!]=?/.test(valStr);
        let matches = false;
        if (hasOp) {
          matches = matchesCriterion(cellVal, valStr);
        } else {
          matches = String(cellVal ?? '').trim().toLocaleLowerCase('tr-TR') === valStr.toLocaleLowerCase('tr-TR');
        }
        if (!matches) {
          rowMatches = false;
          break;
        }
      }
    }

    if (!rowMatches) {
      hidden.push(r);
    }
  }

  return {
    ...sheet,
    hiddenRows: hidden,
    filterConfig: {
      ...config,
      isActive: true,
    },
  };
}

/**
 * Clears all active filters on the worksheet
 */
export function clearSheetFilters(sheet: Worksheet): Worksheet {
  return {
    ...sheet,
    hiddenRows: [],
    filterConfig: sheet.filterConfig
      ? { ...sheet.filterConfig, filters: {}, isActive: false }
      : undefined,
  };
}
