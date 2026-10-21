/**
 * EĞERSAY (COUNTIF) Fonksiyonu Dersi - Excel Akademi AI
 * Aşama 5: Koşullu Sayma ve Veri Analizi
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateFormulaGeneral,
} from '../../modules/learning/taskValidator.ts';

export function createCountifInitialWorkbook(): Workbook {
  return {
    id: 'wb_countif_lesson',
    name: 'Satis_Bolgeleri_Raporu.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Siparişler',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Müşteri', computed: 'Müşteri', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Şehir', computed: 'Şehir', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Sipariş Tutarı (₺)', computed: 'Sipariş Tutarı (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Atlas Ltd.', computed: 'Atlas Ltd.' },
          B2: { raw: 'İstanbul', computed: 'İstanbul' },
          C2: { raw: '1500', computed: 1500, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Birlik A.Ş.', computed: 'Birlik A.Ş.' },
          B3: { raw: 'Ankara', computed: 'Ankara' },
          C3: { raw: '2400', computed: 2400, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Cihan Dış Tic.', computed: 'Cihan Dış Tic.' },
          B4: { raw: 'İstanbul', computed: 'İstanbul' },
          C4: { raw: '850', computed: 850, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'Doğan Lojistik', computed: 'Doğan Lojistik' },
          B5: { raw: 'İzmir', computed: 'İzmir' },
          C5: { raw: '3200', computed: 3200, style: { align: 'right', format: 'currency' } },
          A6: { raw: 'Ege Dağıtım', computed: 'Ege Dağıtım' },
          B6: { raw: 'İstanbul', computed: 'İstanbul' },
          C6: { raw: '1200', computed: 1200, style: { align: 'right', format: 'currency' } },
          A8: { raw: 'İstanbul Sipariş Sayısı:', computed: 'İstanbul Sipariş Sayısı:', style: { bold: true } },
          B8: { raw: '', computed: null, style: { bold: true, align: 'right', bgColor: '#e2e8f0' } },
          A9: { raw: '1000 TL Üstü Siparişler:', computed: '1000 TL Üstü Siparişler:', style: { bold: true } },
          B9: { raw: '', computed: null, style: { bold: true, align: 'right', bgColor: '#e2e8f0' } },
        },
      },
    ],
  };
}

export function createCountifPracticeWorkbook(): Workbook {
  return {
    id: 'wb_countif_practice',
    name: 'Depo_Stok_Sayimi.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Stok Durumu',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Ürün', computed: 'Ürün', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Stok Adedi', computed: 'Stok Adedi', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Monitör 24"', computed: 'Monitör 24"' },
          B2: { raw: '75', computed: 75, style: { align: 'right' } },
          A3: { raw: 'Kablosuz Klavye', computed: 'Kablosuz Klavye' },
          B3: { raw: '30', computed: 30, style: { align: 'right' } },
          A4: { raw: 'Optik Mouse', computed: 'Optik Mouse' },
          B4: { raw: '120', computed: 120, style: { align: 'right' } },
          A5: { raw: 'USB Hub', computed: 'USB Hub' },
          B5: { raw: '45', computed: 45, style: { align: 'right' } },
          A7: { raw: 'Kritik Seviye Üstü (>=50):', computed: 'Kritik Seviye Üstü (>=50):', style: { bold: true } },
          B7: { raw: '', computed: null, style: { bold: true, align: 'right', bgColor: '#e2e8f0' } },
        },
      },
    ],
  };
}

export const countifFunctionLesson: Lesson = {
  id: 'countif_function',
  number: 4,
  title: 'EĞERSAY (COUNTIF) ile Koşullu Sayma',
  lessonType: 'formula_function',
  category: 'İstatistiksel & Koşullu Formüller',
  level: 'Temel-Orta',
  estimatedMinutes: 14,
  description:
    'Belirli bir şarta veya ölçüte uyan hücrelerin sayısını tek bir formülle dinamik olarak saymayı öğrenin.',
  learningObjective:
    'EĞERSAY (COUNTIF) fonksiyonunun sözdizimini, aralık ve ölçüt parametrelerini, metin ve sayısal karşılaştırma kurallarını kavramak.',
  learningOutcomes: [
    'Belirli bir metni içeren hücrelerin sayısını hatasız bulabilme',
    'Sayısal ölçütleri (">1000", ">=50") çift tırnak içinde doğru biçimlendirebilme',
    'EĞERSAY ve COUNTIF fonksiyonlarını çalışma tablolarında pratik olarak uygulayabilme',
  ],
  prerequisites: [
    {
      id: 'sum_function',
      title: 'TOPLA (SUM) Fonksiyonu',
    },
    {
      id: 'if_function',
      title: 'EĞER (IF) Fonksiyonu',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde büyük veri listelerinde sadece istediğimiz kritere uyan satırları saymamızı sağlayan EĞERSAY fonksiyonunu inceleyeceğiz. Örneğin bir sipariş listesinde kaç tane "İstanbul" siparişi olduğunu tek bir formülle saniyeler içinde hesaplayacağız.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'EĞERSAY Fonksiyonunun Amacı',
      content:
        'EĞERSAY fonksiyonu, belirtilen bir hücre aralığında yalnızca belirlediğimiz bir ölçütle (kriterle) eşleşen hücreleri sayar. Normal BAĞ_DEĞ_SAY boş olmayan sayıları sayarken, EĞERSAY "sadece İstanbul olanlar" veya "sadece 1000\'den büyük olanlar" gibi filtrelenmiş sayım yapar.',
      keyTakeaway: 'Belirli bir kritere uyan hücrelerin adedini sayar.',
      audioNarration:
        'EĞERSAY fonksiyonu, belirlediğimiz bir kurala uyan hücreleri tek tek sayar. Örneğin satış listesinde kaç kere İstanbul geçtiğini otomatik hesaplar.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Sözdizimi ve Parametreler',
      content:
        '=EĞERSAY(aralık; ölçüt)\n\n• aralık: Sayılacak hücrelerin bulunduğu alan (örneğin B2:B6).\n• ölçüt: Eşleşmesini istediğimiz metin, sayı veya mantıksal ifade ("İstanbul" veya ">=1000").\nİngilizce Excel karşılığı =COUNTIF(range, criteria) şeklindedir.',
      keyTakeaway: '2 parametre: Sayılacak hücre aralığı ve aranacak ölçüt.',
      audioNarration:
        'Formülde iki parametre vardır: İlki bakılacak hücre aralığı, ikincisi ise aradığımız ölçüttür. Aralarına noktalı virgül koyarız.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Sayısal ve Metin Ölçütlerinin Yazımı',
      content:
        'Metin ararken ölçüt çift tırnak içinde yazılır: "İstanbul".\nSayısal karşılaştırmalarda ise büyüktür veya küçüktür sembolü yine çift tırnak içine alınır: ">=1000". Doğrudan bir sayıyı arıyorsanız tırnaksız da yazabilirsiniz.',
      keyTakeaway: 'Karşılaştırma işleçleri (>=, <, >) çift tırnak içine yazılır: ">=1000".',
      audioNarration:
        'Ölçüt yazarken en önemli kural şudur: Büyük eşit veya küçük eşit gibi sembolleri kullanırken çift tırnak içine almamız gerekir. Örneğin tırnak içinde büyüktür eşittir bin yazarız.',
    },
  ],

  solvedExample: {
    title: 'Şehir Bazlı Sipariş Sayısı Tespiti',
    scenario:
      'B2:B6 aralığındaki şehir listesinde kaç adet "İstanbul" siparişi olduğunu B8 hücresinde hesaplamak istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'B8 hücresine tıklayın.',
        tip: 'Sonucun yazılacağı B8 hücresini seçin.',
      },
      {
        stepNumber: 2,
        instruction: '=EĞERSAY(B2:B6;"İstanbul") yazıp Enter tuşuna basın.',
        tip: 'İngilizce Excel için =COUNTIF(B2:B6;"İstanbul") kullanabilirsiniz.',
      },
    ],
    explanation:
      'Excel B2:B6 aralığındaki 5 satırı tarar; B2, B4 ve B6 hücrelerinde "İstanbul" bulur ve sonuç olarak 3 değerini üretir.',
    audioNarration:
      'Formülümüz B2 ile B6 aralığını tarayarak İstanbul yazan üç satırı tespit eder ve sonucu üç olarak görüntüler.',
  },

  realWorldScenarios: [
    {
      industry: 'E-Ticaret & Lojistik',
      role: 'Lojistik Koordinatörü',
      challenge: 'Günde 2.000 teslimat yapılan bir depoda teslimat durumu "Teslim Edildi" olan siparişlerin sayısını anlık izlemek.',
      solution: '=EĞERSAY(D2:D2000; "Teslim Edildi") formülü ile pano göstergesi oluşturmak.',
      appliedBenefit: 'Depo yöneticisi gün içinde tamamlanan sipariş hacmini tek bakışta gerçek zamanlı takip edebildi.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Sayısal ölçütlerde tırnak işaretini unutmak',
      cause: '=EĞERSAY(C2:C6; >=1000) yazıldığında Excel formül hatası verir.',
      solution: 'Karşılaştırma ifadesini tırnak içine alın: ">=1000".',
    },
    {
      mistake: 'Aralık yerine tek bir hücre belirtmek',
      cause: '=EĞERSAY(B2; "İstanbul") yazıldığında yalnızca tek hücre kontrol edilir.',
      solution: 'Sayım yapmak istediğiniz tüm sütun alanını iki nokta ile belirtin: B2:B6.',
    },
  ],

  tasks: [
    {
      id: 'task_countif_1',
      number: 1,
      title: '1. Adım: B8 Hücresini Seçin',
      instruction: 'İstanbul sipariş adedini yazdıracağımız B8 hücresine tıklayın.',
      narrationText: 'İstanbul sipariş toplamını hesaplayacağımız B8 hücresini seçin.',
      targetCell: 'B8',
      hints: {
        level1: 'Tablonun altındaki "İstanbul Sipariş Sayısı" etiketinin yanındaki hücreye tıklayın.',
        level2: 'B sütunu 8. satırdaki B8 hücresine tıklayın.',
        level3: 'Doğrudan B8 hücresine tıklayın.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'B', 8, 'İstanbul Sipariş Sayısı'),
    },
    {
      id: 'task_countif_2',
      number: 2,
      title: '2. Adım: İstanbul Siparişlerini Sayın',
      instruction:
        'B8 hücresine =EĞERSAY(B2:B6;"İstanbul") yazıp Enter tuşuna basın.',
      narrationText:
        'EĞERSAY formülü ile B2:B6 aralığındaki İstanbul kayıtlarını sayacağız. Sonucun 3 çıkması beklenir.',
      targetCell: 'B8',
      hints: {
        level1: '=EĞERSAY(B2:B6;"İstanbul") formülünü yazın.',
        level2: 'Şehir adını çift tırnak içine almayı unutmayın.',
        level3: 'B8 hücresine =EĞERSAY(B2:B6;"İstanbul") veya =COUNTIF(B2:B6;"İstanbul") yazıp Enter\'a basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'B8',
          [
            '=EĞERSAY(B2:B6;"İstanbul")',
            '=COUNTIF(B2:B6;"İstanbul")',
            '=EĞERSAY(B2:B6,"İstanbul")',
            '=COUNTIF(B2:B6,"İstanbul")',
          ],
          3
        ),
    },
    {
      id: 'task_countif_3',
      number: 3,
      title: '3. Adım: 1000 TL ve Üzeri Siparişleri Sayın',
      instruction:
        'B9 hücresine tıklayın ve =EĞERSAY(C2:C6;">=1000") yazıp Enter tuşuna basın.',
      narrationText:
        'Bu kez tutar sütunundaki (C2:C6) 1000 TL ve üzerindeki siparişleri sayıyoruz. Sonuç 4 olmalıdır.',
      targetCell: 'B9',
      hints: {
        level1: 'B9 hücresini seçin ve C2:C6 aralığını ">=1000" ölçütü ile saydırın.',
        level2: '=EĞERSAY(C2:C6;">=1000") yazın.',
        level3: '=EĞERSAY(C2:C6;">=1000") veya =COUNTIF(C2:C6;">=1000") yazıp Enter tuşuna basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'B9',
          [
            '=EĞERSAY(C2:C6;">=1000")',
            '=COUNTIF(C2:C6;">=1000")',
            '=EĞERSAY(C2:C6,">=1000")',
            '=COUNTIF(C2:C6,">=1000")',
          ],
          4
        ),
    },
  ],

  independentPractice: {
    goal: 'Depo stok listesinde stok adedi 50 ve üzerinde olan kritik seviye üstü ürün sayısını B7 hücresinde hesaplayın.',
    scenario:
      'Depo sorumlusu olarak B2:B5 aralığındaki stokları denetleyin. Stok adedi 50\'den büyük veya eşit olan ürünlerin sayısını EĞERSAY formülü ile bulun.',
    narrationText:
      'Bağımsız alıştırmada stok adetleri 50 ve üzerinde olan ürünleri EĞERSAY formülü ile sayın.',
    workbookFactory: createCountifPracticeWorkbook,
    targetCell: 'B7',
    hints: [
      'B7 hücresini seçin.',
      'Aralık B2:B5, ölçüt ise ">=50" olmalıdır.',
      '=EĞERSAY(B2:B5;">=50") yazıp Enter\'a basın. Beklenen sonuç: 2 (75 ve 120).',
    ],
    validate: (sheet) =>
      validateFormulaGeneral(
        sheet,
        'B7',
        [
          '=EĞERSAY(B2:B5;">=50")',
          '=COUNTIF(B2:B5;">=50")',
          '=EĞERSAY(B2:B5,">=50")',
          '=COUNTIF(B2:B5,">=50")',
        ],
        2
      ),
  },

  quizQuestions: [
    {
      id: 'q_countif_1',
      question: '=EĞERSAY(A1:A10; ">50") formülü ne anlama gelir?',
      options: [
        'A1 ile A10 arasındaki hücrelerden değeri 50\'den büyük olanların sayısını verir',
        'A1 ile A10 arasındaki 50\'den büyük sayıları toplar',
        'A50 hücresinin değerini verir',
        'İlk 50 hücreyi sayar',
      ],
      correctIndex: 0,
      explanation: 'EĞERSAY belirtilen ölçüte uyan hücre adedini sayar, toplamaz.',
    },
    {
      id: 'q_countif_2',
      question: 'EĞERSAY fonksiyonunda İngilizce karşılık aşağıdakilerden hangisidir?',
      options: ['COUNTIF', 'SUMIF', 'COUNTA', 'IFCOUNT'],
      correctIndex: 0,
      explanation: 'EĞERSAY fonksiyonunun İngilizce karşılığı COUNTIF\'tir.',
    },
  ],

  initialWorkbookFactory: createCountifInitialWorkbook,
};
