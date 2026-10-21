/**
 * Task validation service that inspects live spreadsheet state
 * Deterministic rules with contextual diagnostic feedback
 * Stage 4: Expanded with Cell Formatting, Alignment, Range and Interface validation
 */

import type { Worksheet, CellCoordinate, SelectionRange, NumberFormat, HorizontalAlignment, CellStyle } from '../spreadsheet/types.ts';
import type { TaskValidationResult } from './lessonTypes.ts';
import { normalizeFormula, parseNumeric, isNumeric } from '../../utils/validation.ts';
import { parseRange, getAddressesInRange } from '../../utils/cellAddress.ts';

export function validateCellSelection(
  activeCell: CellCoordinate,
  expectedCol: string,
  expectedRow: number,
  cellName: string = 'Toplam hücresi'
): TaskValidationResult {
  const isMatch =
    activeCell.col.toUpperCase() === expectedCol.toUpperCase() &&
    activeCell.row === expectedRow;

  if (isMatch) {
    return {
      isValid: true,
      message: `Harika! ${expectedCol}${expectedRow} (${cellName}) hücresini başarıyla seçtiniz.`,
      expectedValue: `${expectedCol}${expectedRow}`,
    };
  }

  // Diagnostic feedback if user clicked nearby cells
  let diagnostic = `Şu anda ${activeCell.col}${activeCell.row} hücresindesiniz. Lütfen ${expectedCol}${expectedRow} hücresine tıklayın.`;
  if (activeCell.col.toUpperCase() === 'B' && activeCell.row === 6) {
    diagnostic =
      'Toplam hücresini seçmek istiyorduk. Şu anda B6 hücresi seçili. B6, son ürünün (Kulaklık) satış tutarını içeriyor. Bir alt satırdaki B7 hücresini seçmeyi deneyin.';
  } else if (activeCell.col.toUpperCase() === 'A' && activeCell.row === 7) {
    diagnostic =
      'A7 hücresinde "Toplam" metin etiketi bulunuyor. Sayısal toplamı yazacağımız hücre hemen yanındaki B7 hücresidir.';
  }

  return {
    isValid: false,
    message: diagnostic,
    hint: `${expectedCol}${expectedRow} hücresi tablonun toplam satırındaki boş kutudur. Ad kutusuna da yazabilirsiniz.`,
    expectedValue: `${expectedCol}${expectedRow}`,
  };
}

export function validateFormulaBarOrEditFocus(
  activeCell: CellCoordinate,
  expectedCol: string,
  expectedRow: number,
  isEditing?: boolean
): TaskValidationResult {
  const isCellMatch =
    activeCell.col.toUpperCase() === expectedCol.toUpperCase() &&
    activeCell.row === expectedRow;

  if (!isCellMatch) {
    return {
      isValid: false,
      message: `Önce ${expectedCol}${expectedRow} hücresini seçmelisiniz. Şu an ${activeCell.col}${activeCell.row} hücresi seçili.`,
      hint: `${expectedCol}${expectedRow} hücresine tıklayın.`,
    };
  }

  if (isEditing) {
    return {
      isValid: true,
      message: `Harika! Formül giriş alanı düzenleme modunda etkinleştirildi.`,
    };
  }

  return {
    isValid: false,
    message: `${expectedCol}${expectedRow} hücresi seçili, ancak formül giriş alanı henüz düzenleme modunda değil.`,
    hint: `Üstteki fx Formül Çubuğuna tıklayın veya klavyenizden F2 tuşuna basın.`,
  };
}

export function validateFormulaInput(
  sheet: Worksheet,
  targetCellAddr: string,
  acceptedFormulas: string[]
): TaskValidationResult {
  const cell = sheet.cells[targetCellAddr];
  if (!cell || !cell.raw || cell.raw.trim() === '') {
    return {
      isValid: false,
      message: `${targetCellAddr} hücresine henüz bir formül girilmedi.`,
      hint: `Formül çubuğuna veya ${targetCellAddr} hücresine =TOPLA(B2:B6) yazıp Enter'a basın.`,
    };
  }

  const rawTrimmed = cell.raw.trim();

  // Diagnostic: Missing equals sign
  if (!rawTrimmed.startsWith('=')) {
    return {
      isValid: false,
      message: `Formülün başına eşittir (=) işareti koymalısınız. '${rawTrimmed}' şu an düz metin olarak algılanıyor.`,
      hint: `Excel'de tüm formüller = işareti ile başlar. Örnek: =TOPLA(B2:B6)`,
      expectedValue: acceptedFormulas[0],
    };
  }

  const normalizedRaw = normalizeFormula(cell.raw);
  const normalizedAccepted = acceptedFormulas.map(normalizeFormula);

  // Diagnostic: Semicolon used instead of colon in range
  if (normalizedRaw.includes('B2;B6')) {
    return {
      isValid: false,
      message:
        'Aralık belirtirken iki nokta (:) kullanılmalıdır. Noktalı virgül (;) yalnızca B2 ve B6 hücrelerini toplar, aradaki değerleri toplamaz!',
      hint: '=TOPLA(B2:B6) şeklinde iki nokta kullanın.',
      expectedValue: acceptedFormulas[0],
    };
  }

  const isFormulaMatch = normalizedAccepted.includes(normalizedRaw);

  if (isFormulaMatch) {
    return {
      isValid: true,
      message: `Tebrikler! ${targetCellAddr} hücresine doğru formül yazıldı: ${cell.raw}`,
      expectedValue: cell.raw,
    };
  }

  return {
    isValid: false,
    message: `${targetCellAddr} hücresindeki formül '${cell.raw}' beklenen formül ile eşleşmiyor.`,
    hint: `Beklenen formül: =TOPLA(B2:B6) veya =SUM(B2:B6). Büyük/küçük harf fark etmez.`,
    expectedValue: acceptedFormulas[0],
  };
}

export function validateCellNumericResult(
  sheet: Worksheet,
  targetCellAddr: string,
  expectedSum: number
): TaskValidationResult {
  const cell = sheet.cells[targetCellAddr];
  if (!cell || cell.computed === null || cell.computed === undefined || cell.computed === '') {
    return {
      isValid: false,
      message: `${targetCellAddr} hücresinde henüz bir sonuç hesaplanmadı. Formülü yazıp Enter tuşuna basın.`,
      hint: `Enter tuşuna basarak formül girişini onaylayın.`,
    };
  }

  const computedNum = parseNumeric(cell.computed);

  if (Math.abs(computedNum - expectedSum) < 0.001) {
    return {
      isValid: true,
      message: `Mükemmel! Toplam sonuç tam olarak ${expectedSum} ₺ olarak hesaplandı.`,
      expectedValue: expectedSum,
    };
  }

  return {
    isValid: false,
    message: `Hesaplanan sonuç (${cell.computed}) beklenen değer (${expectedSum}) ile uyuşmuyor.`,
    hint: `Lütfen toplanan hücre aralığının B2:B6 olduğundan emin olun (350+450+1200+650+400 = 3050).`,
    expectedValue: expectedSum,
  };
}

export function validateAutoRecalculation(
  sheet: Worksheet,
  modifiedCellAddr: string,
  expectedModifiedVal: number,
  targetSumAddr: string,
  expectedNewSum: number
): TaskValidationResult {
  const modifiedCell = sheet.cells[modifiedCellAddr];
  const targetCell = sheet.cells[targetSumAddr];

  const currentModVal = modifiedCell ? parseNumeric(modifiedCell.computed) : 0;
  const currentSumVal = targetCell ? parseNumeric(targetCell.computed) : 0;

  if (Math.abs(currentModVal - expectedModifiedVal) > 0.001) {
    return {
      isValid: false,
      message: `${modifiedCellAddr} hücresinin değeri henüz ${expectedModifiedVal} yapılmadı. Şu anki değer: ${currentModVal}.`,
      hint: `${modifiedCellAddr} hücresine tıklayın, 500 yazıp Enter'a basın.`,
      expectedValue: expectedModifiedVal,
    };
  }

  if (Math.abs(currentSumVal - expectedNewSum) < 0.001) {
    return {
      isValid: true,
      message: `Harika! B3 değeri 500 olunca, B7 toplamı otomatik olarak ${expectedNewSum} ₺ oldu!`,
      expectedValue: expectedNewSum,
    };
  }

  return {
    isValid: false,
    message: `${targetSumAddr} hücresindeki toplam ${expectedNewSum} yerine ${currentSumVal} olarak görünüyor.`,
    hint: `B7 hücresindeki formülün =TOPLA(B2:B6) olduğundan emin olun.`,
    expectedValue: expectedNewSum,
  };
}

/**
 * Stage 4 Universal Validators:
 * Cell Formatting, Alignment, Currency, Borders, Background Color, Range Selection
 */

export function resolveTargetAddresses(targetAddrOrRange: string): string[] {
  if (targetAddrOrRange.includes(':')) {
    const range = parseRange(targetAddrOrRange);
    return range ? getAddressesInRange(range) : [targetAddrOrRange];
  }
  return [targetAddrOrRange];
}

export function validateCellFormatting(
  sheet: Worksheet,
  targetCellOrRange: string,
  expectedStyle: {
    bold?: boolean;
    italic?: boolean;
    format?: NumberFormat;
    align?: HorizontalAlignment;
    backgroundColor?: string;
  },
  styleLabel: string = 'İstenen biçimlendirme'
): TaskValidationResult {
  const addresses = resolveTargetAddresses(targetCellOrRange);

  if (addresses.length === 0) {
    return {
      isValid: false,
      message: `Geçersiz hedef adres: ${targetCellOrRange}`,
    };
  }

  for (const addr of addresses) {
    const cell = sheet.cells[addr];
    const style: CellStyle = cell?.style || {};

    if (expectedStyle.bold !== undefined && style.bold !== expectedStyle.bold) {
      return {
        isValid: false,
        message: `${addr} hücresi henüz Kalın (Bold) yapılmadı. Lütfen araç çubuğundaki B (Kalın) düğmesine tıklayın veya Ctrl+B kısayolunu kullanın.`,
        hint: `${targetCellOrRange} aralığını seçin ve araç çubuğundaki Kalın (B) simgesine tıklayın.`,
      };
    }

    if (expectedStyle.italic !== undefined && style.italic !== expectedStyle.italic) {
      return {
        isValid: false,
        message: `${addr} hücresi henüz İtalik yapılmadı. Araç çubuğundaki I (İtalik) düğmesine tıklayın.`,
        hint: `İtalik yapmak için araç çubuğundaki I simgesine veya Ctrl+I kısayoluna basın.`,
      };
    }

    if (expectedStyle.format !== undefined && style.format !== expectedStyle.format) {
      const formatNames: Record<string, string> = {
        currency: 'Finansal Para Birimi (₺)',
        percentage: 'Yüzde (%)',
        number: 'Sayı',
        general: 'Genel',
      };
      const expectedName = formatNames[expectedStyle.format] || expectedStyle.format;
      return {
        isValid: false,
        message: `${addr} hücresine henüz ${expectedName} biçimlendirmesi uygulanmadı.`,
        hint: `${targetCellOrRange} hücrelerini seçtikten sonra araç çubuğundaki ₺ (Para Birimi) butonuna tıklayın.`,
      };
    }

    if (expectedStyle.align !== undefined && style.align !== expectedStyle.align) {
      const alignNames: Record<string, string> = {
        center: 'Ortala',
        left: 'Sola Hizala',
        right: 'Sağa Hizala',
      };
      return {
        isValid: false,
        message: `${addr} hücresi ${alignNames[expectedStyle.align] || expectedStyle.align} olarak hizalanmadı.`,
        hint: `Araç çubuğundaki hizalama düğmelerinden Ortala butonunu seçin.`,
      };
    }

    if (expectedStyle.backgroundColor && style.backgroundColor !== expectedStyle.backgroundColor) {
      return {
        isValid: false,
        message: `${addr} hücresine dolgu rengi uygulanmadı.`,
        hint: `Dolgu rengi paletinden belirtilen rengi seçin.`,
      };
    }
  }

  return {
    isValid: true,
    message: `Tebrikler! ${targetCellOrRange} alanına ${styleLabel} başarıyla uygulandı.`,
  };
}

export function validateRangeSelection(
  selectionRange: SelectionRange | null,
  activeCell: CellCoordinate,
  expectedRangeStr: string
): TaskValidationResult {
  const expected = parseRange(expectedRangeStr);
  if (!expected) {
    return {
      isValid: false,
      message: `Geçersiz hedef aralık tanımı: ${expectedRangeStr}`,
    };
  }

  const expectedAddresses = getAddressesInRange(expected);

  if (!selectionRange) {
    const isSingleCellMatch =
      expectedAddresses.length === 1 &&
      activeCell.col.toUpperCase() === expected.startCol.toUpperCase() &&
      activeCell.row === expected.startRow;

    if (isSingleCellMatch) {
      return {
        isValid: true,
        message: `Harika! ${expectedRangeStr} hücresi seçildi.`,
      };
    }

    return {
      isValid: false,
      message: `Lütfen ${expectedRangeStr} aralığını seçin. Şu anda yalnızca ${activeCell.col}${activeCell.row} hücresi seçili.`,
      hint: `Shift tuşuna basılı tutarak veya fare ile sürükleyerek ${expectedRangeStr} aralığını seçebilirsiniz.`,
    };
  }

  const isColMatch =
    selectionRange.startCol.toUpperCase() === expected.startCol.toUpperCase() &&
    selectionRange.endCol.toUpperCase() === expected.endCol.toUpperCase();
  const isRowMatch =
    selectionRange.startRow === expected.startRow &&
    selectionRange.endRow === expected.endRow;

  if (isColMatch && isRowMatch) {
    return {
      isValid: true,
      message: `Harika! ${expectedRangeStr} aralığı tam olarak seçildi.`,
    };
  }

  return {
    isValid: false,
    message: `Seçili aralık (${selectionRange.startCol}${selectionRange.startRow}:${selectionRange.endCol}${selectionRange.endRow}) beklenen ${expectedRangeStr} ile eşleşmiyor.`,
    hint: `Farenizi ${expected.startCol}${expected.startRow} hücresinden ${expected.endCol}${expected.endRow} hücresine kadar sürükleyin.`,
  };
}

/**
 * Stage 5: Validates formula input and computed result for any general formula
 */
export function validateFormulaGeneral(
  sheet: Worksheet,
  targetCellAddr: string,
  acceptedFormulas: string[],
  expectedComputed?: string | number
): TaskValidationResult {
  const cell = sheet.cells[targetCellAddr];
  if (!cell || !cell.raw || cell.raw.trim() === '') {
    return {
      isValid: false,
      message: `${targetCellAddr} hücresine henüz bir formül girilmedi.`,
      hint: `Formül çubuğuna tıklayın veya ${targetCellAddr} hücresine formülü yazıp Enter'a basın.`,
    };
  }

  const rawTrimmed = cell.raw.trim();
  if (!rawTrimmed.startsWith('=')) {
    return {
      isValid: false,
      message: `Formülün başına eşittir (=) işareti koymalısınız. '${rawTrimmed}' şu an düz metin olarak algılanıyor.`,
      hint: `Excel formülleri her zaman = işareti ile başlar.`,
      expectedValue: acceptedFormulas[0],
    };
  }

  const normalizedRaw = normalizeFormula(cell.raw);
  const normalizedAccepted = acceptedFormulas.map(normalizeFormula);
  const isFormulaMatch = normalizedAccepted.includes(normalizedRaw);

  if (!isFormulaMatch) {
    // If user got the right computed result but slight syntax deviation, allow with positive feedback
    if (expectedComputed !== undefined && cell.computed !== undefined && cell.computed !== null) {
      const computedMatches =
        typeof expectedComputed === 'number'
          ? Math.abs(parseNumeric(cell.computed) - expectedComputed) < 0.001
          : String(cell.computed).trim().toLowerCase() === String(expectedComputed).trim().toLowerCase();

      if (computedMatches) {
        return {
          isValid: true,
          message: `Tebrikler! ${targetCellAddr} hücresinde doğru sonuç elde edildi: ${cell.computed} (${cell.raw})`,
          expectedValue: cell.raw,
        };
      }
    }

    return {
      isValid: false,
      message: `${targetCellAddr} hücresindeki formül '${cell.raw}' beklenen formül ile uyuşmuyor.`,
      hint: `Örnek kullanım: ${acceptedFormulas[0]}`,
      expectedValue: acceptedFormulas[0],
    };
  }

  // Formula matched - also check computed value if specified
  if (expectedComputed !== undefined) {
    if (cell.computed === null || cell.computed === undefined || cell.computed === '') {
      return {
        isValid: false,
        message: `${targetCellAddr} hücresinde formül yazıldı ancak Enter tuşuna basılarak hesaplama tamamlanmadı.`,
        hint: `Enter tuşuna basarak formülü onaylayın.`,
      };
    }

    const matches =
      typeof expectedComputed === 'number'
        ? Math.abs(parseNumeric(cell.computed) - expectedComputed) < 0.001
        : String(cell.computed).trim().toLowerCase() === String(expectedComputed).trim().toLowerCase();

    if (!matches) {
      return {
        isValid: false,
        message: `${targetCellAddr} hücresinde hesaplanan değer '${cell.computed}', beklenen '${expectedComputed}' değeriyle uyuşmuyor.`,
        hint: `Formül argümanlarını ve hücre başvurularını kontrol edin: ${acceptedFormulas[0]}`,
        expectedValue: String(expectedComputed),
      };
    }
  }

  return {
    isValid: true,
    message: `Tebrikler! ${targetCellAddr} hücresine doğru formül uygulandı ve sonuç başarıyla hesaplandı: ${cell.computed ?? cell.raw}`,
    expectedValue: cell.raw,
  };
}

/**
 * Stage 5: Validates table sorting order on target column
 */
export function validateTableSorted(
  sheet: Worksheet,
  targetCol: string,
  ascending: boolean,
  bounds: { startRow: number; endRow: number }
): TaskValidationResult {
  const col = targetCol.toUpperCase();
  const values: (string | number)[] = [];

  for (let r = bounds.startRow; r <= bounds.endRow; r++) {
    const cell = sheet.cells[`${col}${r}`];
    const val = cell ? (cell.computed !== null && cell.computed !== undefined ? cell.computed : cell.raw) : '';
    values.push(val);
  }

  // Check if sorted
  let isSorted = true;
  for (let i = 0; i < values.length - 1; i++) {
    const a = values[i];
    const b = values[i + 1];

    const bothNum = isNumeric(a) && isNumeric(b) && a !== '' && b !== '';

    let cmp = 0;
    if (bothNum) {
      cmp = parseNumeric(a) - parseNumeric(b);
    } else {
      cmp = String(a).localeCompare(String(b), 'tr', { sensitivity: 'base' });
    }

    if (ascending && cmp > 0) {
      isSorted = false;
      break;
    }
    if (!ascending && cmp < 0) {
      isSorted = false;
      break;
    }
  }

  if (isSorted) {
    return {
      isValid: true,
      message: `Harika! ${col} sütunu ${ascending ? 'A-Z (Küçükten Büyüğe)' : 'Z-A (Büyükten Küçüğe)'} başarıyla sıralandı.`,
    };
  }

  return {
    isValid: false,
    message: `${col} sütunundaki veriler henüz beklenen sırada (${ascending ? 'A-Z' : 'Z-A'}) değil.`,
    hint: `Üst şeritteki Sırala butonunu (veya sütun başlığındaki açılır menüyü) kullanarak ${col} sütununu ${ascending ? 'A\'dan Z\'ye' : 'Z\'den A\'ya'} sıralayın.`,
  };
}

/**
 * Stage 5: Validates filter application
 */
export function validateTableFiltered(
  sheet: Worksheet,
  expectedFilterCol?: string,
  expectedHiddenMinCount?: number
): TaskValidationResult {
  const filterConfig = sheet.filterConfig;
  if (!filterConfig || !filterConfig.isActive) {
    return {
      isValid: false,
      message: `Çalışma sayfasında filtreleme henüz aktif edilmedi.`,
      hint: `Araç çubuğundaki 'Filtrele' butonuna tıklayarak tablo başlıklarına filtre butonlarını ekleyin.`,
    };
  }

  if (expectedFilterCol) {
    const colFilter = filterConfig.filters[expectedFilterCol.toUpperCase()];
    if (!colFilter || colFilter.condition === 'none') {
      return {
        isValid: false,
        message: `${expectedFilterCol.toUpperCase()} sütununa henüz bir filtre uygulanmadı.`,
        hint: `${expectedFilterCol.toUpperCase()} sütun başlığındaki filtre simgesine tıklayıp filtre ölçütü belirleyin.`,
      };
    }
  }

  if (expectedHiddenMinCount !== undefined) {
    const hiddenCount = (sheet.hiddenRows || []).length;
    if (hiddenCount < expectedHiddenMinCount) {
      return {
        isValid: false,
        message: `Filtre uygulandı ancak beklenen sayıda satır filtrelenmedi (Gizlenen satır: ${hiddenCount}, Beklenen en az: ${expectedHiddenMinCount}).`,
        hint: `Filtre ölçütünüzü kontrol edin.`,
      };
    }
  }

  return {
    isValid: true,
    message: `Tebrikler! Tablo başarıyla filtrelendi.`,
  };
}

/**
 * Stage 5: Validates chart creation
 */
export function validateChartExists(
  sheet: Worksheet,
  expectedType?: 'column' | 'line' | 'pie',
  expectedValueRange?: string
): TaskValidationResult {
  const charts = sheet.charts || [];
  if (charts.length === 0) {
    return {
      isValid: false,
      message: `Çalışma sayfasında henüz bir grafik oluşturulmadı.`,
      hint: `Araç çubuğundaki 'Grafik Ekle' butonuna tıklayarak yeni bir grafik ekleyin.`,
    };
  }

  if (expectedType) {
    const matchingChart = charts.find((c) => c.type === expectedType);
    if (!matchingChart) {
      return {
        isValid: false,
        message: `Oluşturulan grafik '${expectedType}' türünde değil (Mevcut: ${charts[0].type}).`,
        hint: `Lütfen '${expectedType}' türünde bir grafik oluşturun.`,
      };
    }
  }

  if (expectedValueRange) {
    const matchingRange = charts.find(
      (c) => c.valueRange.toUpperCase() === expectedValueRange.toUpperCase()
    );
    if (!matchingRange) {
      return {
        isValid: false,
        message: `Grafiğin değer aralığı ${expectedValueRange} olmalıdır (Mevcut: ${charts[0].valueRange}).`,
        hint: `Grafik ekleme penceresinde Değer Aralığı olarak ${expectedValueRange} girin.`,
      };
    }
  }

  return {
    isValid: true,
    message: `Harika! Grafik başarıyla oluşturuldu ve çalışma sayfasına eklendi.`,
  };
}

