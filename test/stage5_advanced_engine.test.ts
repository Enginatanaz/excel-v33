/**
 * Automated Test Suite for Stage 5:
 * Formula Engine, Data Management & Chart Simulator
 */

import { evaluateFormula } from '../src/modules/spreadsheet/formulaEngine.ts';
import { sortWorksheet, applySheetFilters, clearSheetFilters } from '../src/modules/spreadsheet/dataManagement.ts';
import { validateLessonQuality } from '../src/modules/learning/lessonQualityEngine.ts';
import {
  ifFunctionLesson,
  countifFunctionLesson,
  sumifsFunctionLesson,
  vlookupFunctionLesson,
  xlookupFunctionLesson,
  dataSortLesson,
  dataFilterLesson,
  chartCreationLesson,
} from '../src/data/lessons/index.ts';
import type { Worksheet } from '../src/modules/spreadsheet/types.ts';

let totalTests = 0;
let passedTests = 0;

function assert(condition: boolean, testName: string, details?: string) {
  totalTests++;
  if (condition) {
    passedTests++;
    console.log(`  ✓ ${testName}`);
  } else {
    console.error(`  ✗ FAIL: ${testName} ${details ? `(${details})` : ''}`);
  }
}

console.log('\n======================================================');
console.log('🧪 STAGE 5 TEST SUITE: FORMULA, DATA & CHART SIMULATOR');
console.log('======================================================\n');

// ---------------------------------------------------------------------
// 1. FORMULA ENGINE EXPANSION TESTS
// ---------------------------------------------------------------------
console.log('--- 1. Formula Engine Expansion ---');

const sampleSheet: Worksheet = {
  id: 'sheet1',
  name: 'TestSheet',
  rowCount: 10,
  colCount: 6,
  cells: {
    A1: { raw: 'Ad', computed: 'Ad' },
    B1: { raw: 'Şehir', computed: 'Şehir' },
    C1: { raw: 'Not', computed: 'Not' },
    D1: { raw: 'Maaş', computed: 'Maaş' },
    A2: { raw: 'Ahmet', computed: 'Ahmet' },
    B2: { raw: 'İstanbul', computed: 'İstanbul' },
    C2: { raw: '75', computed: 75 },
    D2: { raw: '15000', computed: 15000 },
    A3: { raw: 'Zeynep', computed: 'Zeynep' },
    B3: { raw: 'Ankara', computed: 'Ankara' },
    C3: { raw: '42', computed: 42 },
    D3: { raw: '22000', computed: 22000 },
    A4: { raw: 'Mehmet', computed: 'Mehmet' },
    B4: { raw: 'İstanbul', computed: 'İstanbul' },
    C4: { raw: '88', computed: 88 },
    D4: { raw: '18000', computed: 18000 },
  },
};

// A. EĞER (IF)
const ifResult1 = evaluateFormula('=EĞER(C2>=50;"Geçti";"Kaldı")', sampleSheet);
assert(ifResult1.value === 'Geçti', 'EĞER: Doğru koşul "Geçti" döndürmeli');

const ifResult2 = evaluateFormula('=IF(C3>=50;"Geçti";"Kaldı")', sampleSheet);
assert(ifResult2.value === 'Kaldı', 'IF: İngilizce ad ve yanlış koşul "Kaldı" döndürmeli');

const ifResult3 = evaluateFormula('=EĞER(C2>50; C2*2; 0)', sampleSheet);
assert(ifResult3.value === 150, 'EĞER: Sayısal sonuç ve matematiksel ifade hesaplamalı');

const ifNested = evaluateFormula('=EĞER(C2>80;"A";EĞER(C2>70;"B";"C"))', sampleSheet);
assert(ifNested.value === 'B', 'EĞER: İç içe (nested) EĞER doğru dalı çalıştırmalı');

// B. EĞERSAY (COUNTIF)
const countif1 = evaluateFormula('=EĞERSAY(B2:B4;"İstanbul")', sampleSheet);
assert(countif1.value === 2, 'EĞERSAY: Metin eşleşen hücre sayısını bulmalı (2 İstanbul)');

const countif2 = evaluateFormula('=COUNTIF(C2:C4;">=50")', sampleSheet);
assert(countif2.value === 2, 'COUNTIF: Sayısal karşılaştırma ölçütü ile saymalı (75 ve 88 >= 50)');

const countif3 = evaluateFormula('=EĞERSAY(B2:B4;"İzmir")', sampleSheet);
assert(countif3.value === 0, 'EĞERSAY: Eşleşmeyen aramalarda 0 döndürmeli');

// C. ÇOKETOPLA (SUMIFS)
const sumifs1 = evaluateFormula('=ÇOKETOPLA(D2:D4; B2:B4; "İstanbul"; C2:C4; ">=50")', sampleSheet);
assert(sumifs1.value === 33000, 'ÇOKETOPLA: İki kriterli toplama (15000 + 18000 = 33000)');

const sumifs2 = evaluateFormula('=SUMIFS(D2:D4; B2:B4; "Ankara")', sampleSheet);
assert(sumifs2.value === 22000, 'SUMIFS: İngilizce isim ve tek kriterli ÇOKETOPLA');

const sumifsMismatch = evaluateFormula('=ÇOKETOPLA(D2:D4; B2:B10; "İstanbul")', sampleSheet);
assert(sumifsMismatch.value === '#DEĞER!', 'ÇOKETOPLA: Boyut uyumsuzluğunda #DEĞER! üretmeli');

// D. DÜŞEYARA (VLOOKUP)
const vlookup1 = evaluateFormula('=DÜŞEYARA("Ahmet"; A2:D4; 4; 0)', sampleSheet);
assert(vlookup1.value === 15000, 'DÜŞEYARA: Ahmet için 4. sütundaki Maaşı (15000) getirmeli');

const vlookup2 = evaluateFormula('=VLOOKUP("Zeynep"; A2:D4; 3; 0)', sampleSheet);
assert(vlookup2.value === 42, 'VLOOKUP: İngilizce isim ile 3. sütundaki Notu (42) getirmeli');

const vlookupNotFound = evaluateFormula('=DÜŞEYARA("OlmayanKisi"; A2:D4; 2; 0)', sampleSheet);
assert(vlookupNotFound.value === '#YOK!', 'DÜŞEYARA: Bulunamayan anahtarda #YOK! döndürmeli');

const vlookupBadCol = evaluateFormula('=DÜŞEYARA("Ahmet"; A2:D4; 10; 0)', sampleSheet);
assert(vlookupBadCol.value === '#BAŞV!', 'DÜŞEYARA: Aralık dışı sütun numarasında #BAŞV! üretmeli');

// E. ÇAPRAZARA (XLOOKUP)
const xlookup1 = evaluateFormula('=ÇAPRAZARA("İstanbul"; B2:B4; A2:A4)', sampleSheet);
assert(xlookup1.value === 'Ahmet', 'ÇAPRAZARA: Sol sütundaki değeri getirebilmeli (B -> A)');

const xlookup2 = evaluateFormula('=XLOOKUP("Mehmet"; A2:A4; D2:D4)', sampleSheet);
assert(xlookup2.value === 18000, 'XLOOKUP: İngilizce isim ile değer getirmeli');

const xlookupNotFound = evaluateFormula('=ÇAPRAZARA("Bursa"; B2:B4; A2:A4; "Bulunamadı")', sampleSheet);
assert(xlookupNotFound.value === 'Bulunamadı', 'ÇAPRAZARA: [bulunamıyorsa] özel mesajı döndürmeli');

// ---------------------------------------------------------------------
// 2. DATA MANAGEMENT (SORT & FILTER) TESTS
// ---------------------------------------------------------------------
console.log('\n--- 2. Data Management (Sort & Filter) ---');

const sortSheet: Worksheet = {
  id: 'sort_test',
  name: 'SortTest',
  rowCount: 10,
  colCount: 4,
  cells: {
    A1: { raw: 'İsim', computed: 'İsim' },
    B1: { raw: 'Departman', computed: 'Departman' },
    C1: { raw: 'Maaş', computed: 'Maaş' },
    A2: { raw: 'Zeynep', computed: 'Zeynep' },
    B2: { raw: 'Pazarlama', computed: 'Pazarlama' },
    C2: { raw: '30000', computed: 30000 },
    A3: { raw: 'Ahmet', computed: 'Ahmet' },
    B3: { raw: 'Finans', computed: 'Finans' },
    C3: { raw: '50000', computed: 50000 },
    A4: { raw: 'Burak', computed: 'Burak' },
    B4: { raw: 'Yazılım', computed: 'Yazılım' },
    C4: { raw: '40000', computed: 40000 },
  },
};

// Sort A-Z on Column A
const sortedAsc = sortWorksheet(sortSheet, 'A', true);
assert(sortedAsc.cells['A2'].computed === 'Ahmet', 'Sıralama: A-Z sonrası A2 Ahmet olmalı');
assert(sortedAsc.cells['B2'].computed === 'Finans', 'Satır Bütünlüğü: Ahmet ile Finans birlikte taşınmalı');
assert(sortedAsc.cells['C2'].computed === 50000, 'Satır Bütünlüğü: Ahmet ile 50000 maaş birlikte taşınmalı');
assert(sortedAsc.cells['A4'].computed === 'Zeynep', 'Sıralama: A-Z sonrası A4 Zeynep olmalı');

// Sort Z-A on Column C (Numerical Salary)
const sortedDescSalary = sortWorksheet(sortSheet, 'C', false);
assert(sortedDescSalary.cells['C2'].computed === 50000, 'Sayısal Sıralama: En yüksek maaş (50000) A2/C2 satırında olmalı');
assert(sortedDescSalary.cells['A2'].computed === 'Ahmet', 'Satır Bütünlüğü: 50000 Ahmet ile eşleşmeli');
assert(sortedDescSalary.cells['C4'].computed === 30000, 'Sayısal Sıralama: En düşük maaş (30000) en altta olmalı');

// Filtering
const filtered = applySheetFilters(sortSheet, {
  isActive: true,
  filters: {
    B: { col: 'B', colLetter: 'B', filterType: 'text', condition: 'equals', value: 'Pazarlama' },
  },
});
assert(Boolean(filtered.hiddenRows && filtered.hiddenRows.length === 2), 'Filtreleme: Pazarlama dışındaki 2 satır gizlenmeli');
assert(Boolean(filtered.hiddenRows?.includes(3) && filtered.hiddenRows?.includes(4)), 'Filtreleme: 3 ve 4. satırlar gizli satır listesinde olmalı');

const cleared = clearSheetFilters(filtered);
assert(cleared.hiddenRows?.length === 0, 'Filtre Temizleme: Tüm gizli satırlar geri gelmeli');

// ---------------------------------------------------------------------
// 3. CHART SIMULATOR REACTIVITY
// ---------------------------------------------------------------------
console.log('\n--- 3. Chart Simulator Reactive Extraction ---');

const chartTestSheet: Worksheet = {
  id: 'chart_test',
  name: 'ChartSheet',
  rowCount: 10,
  colCount: 4,
  cells: {
    A1: { raw: 'Ay', computed: 'Ay' },
    B1: { raw: 'Satış', computed: 'Satış' },
    A2: { raw: 'Ocak', computed: 'Ocak' },
    B2: { raw: '1000', computed: 1000 },
    A3: { raw: 'Şubat', computed: 'Şubat' },
    B3: { raw: '2000', computed: 2000 },
  },
  charts: [
    {
      id: 'c1',
      title: 'Aylık Satışlar',
      type: 'column',
      categoryRange: 'A2:A3',
      valueRange: 'B2:B3',
    },
  ],
};

assert(chartTestSheet.charts?.length === 1, 'Grafik: Çalışma sayfasına grafik eklenebilmeli');
assert(chartTestSheet.charts?.[0].type === 'column', 'Grafik: Grafik tipi Sütun olmalı');

// ---------------------------------------------------------------------
// 4. LESSON QUALITY VALIDATION FOR ALL 8 STAGE 5 LESSONS
// ---------------------------------------------------------------------
console.log('\n--- 4. Lesson Quality Verification (100% Target) ---');

const stage5Lessons = [
  ifFunctionLesson,
  countifFunctionLesson,
  sumifsFunctionLesson,
  vlookupFunctionLesson,
  xlookupFunctionLesson,
  dataSortLesson,
  dataFilterLesson,
  chartCreationLesson,
];

for (const lesson of stage5Lessons) {
  const report = validateLessonQuality(lesson);
  const is100 = report.isValid && report.score === 100 && report.errors.length === 0;
  assert(
    is100,
    `Kalite Kontrol (%100): ${lesson.title} [${lesson.lessonType}]`,
    `Skor: ${report.score}, Hatalar: ${report.errors.join('; ')}`
  );
}

console.log('\n======================================================');
console.log(`SUMMARY: ${passedTests} / ${totalTests} tests passed (${Math.round((passedTests / totalTests) * 100)}%)`);
console.log('======================================================\n');

if (passedTests !== totalTests) {
  process.exit(1);
}
