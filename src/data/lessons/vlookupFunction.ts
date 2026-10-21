/**
 * DÜŞEYARA (VLOOKUP) Fonksiyonu Dersi - Excel Akademi AI
 * Aşama 5: Tablo Arama ve Veri Eşleştirme
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateFormulaGeneral,
} from '../../modules/learning/taskValidator.ts';

export function createVlookupInitialWorkbook(): Workbook {
  return {
    id: 'wb_vlookup_lesson',
    name: 'Urun_Katalog_Arama.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Sipariş Sorgu',
        rowCount: 15,
        colCount: 7,
        cells: {
          A1: { raw: 'Aranan Kod', computed: 'Aranan Kod', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          B1: { raw: 'Ürün Adı', computed: 'Ürün Adı', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          C1: { raw: 'Fiyat (₺)', computed: 'Fiyat (₺)', style: { bold: true, align: 'right', bgColor: '#e0f2fe' } },
          A2: { raw: 'U103', computed: 'U103', style: { bold: true } },
          B2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },
          C2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },

          // Master Catalog Table (Cols E to G)
          E1: { raw: 'Ürün Kodu', computed: 'Ürün Kodu', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          F1: { raw: 'Katalog Ürün Adı', computed: 'Katalog Ürün Adı', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          G1: { raw: 'Birim Fiyat (₺)', computed: 'Birim Fiyat (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          E2: { raw: 'U101', computed: 'U101' },
          F2: { raw: 'Kablosuz Klavye', computed: 'Kablosuz Klavye' },
          G2: { raw: '450', computed: 450, style: { align: 'right', format: 'currency' } },
          E3: { raw: 'U102', computed: 'U102' },
          F3: { raw: 'Gaming Mouse', computed: 'Gaming Mouse' },
          G3: { raw: '350', computed: 350, style: { align: 'right', format: 'currency' } },
          E4: { raw: 'U103', computed: 'U103' },
          F4: { raw: '27 inç Monitör', computed: '27 inç Monitör' },
          G4: { raw: '4200', computed: 4200, style: { align: 'right', format: 'currency' } },
          E5: { raw: 'U104', computed: 'U104' },
          F5: { raw: 'Mekanik Klavye', computed: 'Mekanik Klavye' },
          G5: { raw: '1100', computed: 1100, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export function createVlookupPracticeWorkbook(): Workbook {
  return {
    id: 'wb_vlookup_practice',
    name: 'Personel_Departman_Sorgu.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Personel Sorgu',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Sicil No', computed: 'Sicil No', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          B1: { raw: 'Departman', computed: 'Departman', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          A2: { raw: 'P202', computed: 'P202', style: { bold: true } },
          B2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },

          D1: { raw: 'Sicil No', computed: 'Sicil No', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          E1: { raw: 'Departman Adı', computed: 'Departman Adı', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          D2: { raw: 'P201', computed: 'P201' },
          E2: { raw: 'Muhasebe', computed: 'Muhasebe' },
          D3: { raw: 'P202', computed: 'P202' },
          E3: { raw: 'Pazarlama', computed: 'Pazarlama' },
          D4: { raw: 'P203', computed: 'P203' },
          E4: { raw: 'Yazılım', computed: 'Yazılım' },
        },
      },
    ],
  };
}

export const vlookupFunctionLesson: Lesson = {
  id: 'vlookup_function',
  number: 6,
  title: 'DÜŞEYARA (VLOOKUP) ile Veri Eşleştirme',
  lessonType: 'formula_function',
  category: 'Arama ve Başvuru Fonksiyonları',
  level: 'Orta',
  estimatedMinutes: 18,
  description:
    'Farklı tablolardaki verileri ortak bir anahtar değer (ID, barkod, kod) üzerinden eşleştirip getiren efsanevi DÜŞEYARA fonksiyonunu öğrenin.',
  learningObjective:
    'DÜŞEYARA fonksiyonunun 4 parametresini (aranan_değer, tablo_dizisi, sütun_indis_sayısı, aralık_bak) eksiksiz anlamak ve tam eşleşme (0/YANLIŞ) modunu kullanabilmek.',
  learningOutcomes: [
    'Arama anahtarının ana tablonun en sol sütununda bulunması kuralını kavrama',
    'Sütun indis sayısı mantığını öğrenerek tablodan istenen sütunu çekebilme',
    'DÜŞEYARA ve VLOOKUP formüllerini tam eşleşme kuralıyla hatasız çalıştırabilme',
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
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde iş mülakatlarının ve ofis çalışmalarının vazgeçilmezi olan DÜŞEYARA fonksiyonunu öğreneceğiz. Bir ürün kodu girdiğinizde o ürünün adını ve fiyatını ana katalog tablosundan otomatik çekip getirmeyi adım adım uygulayacağız.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'DÜŞEYARA Fonksiyonunun Amacı',
      content:
        'DÜŞEYARA (İngilizce VLOOKUP - Vertical Lookup), büyük bir veri tablosunun en sol sütununda dikey olarak bir değeri arar ve o değerin bulunduğu satırdaki istediğiniz bir sütundaki bilgiyi getirir. Manuel arayıp kopyala yapıştır yapma ihtiyacını tamamen ortadan kaldırır.',
      keyTakeaway: 'Tablonun en sol sütununda arama yaparak istenen sütundaki veriyi getirir.',
      audioNarration:
        'DÜŞEYARA fonksiyonu, dikey bir tablonun en solundaki sütunda arama yapar. Eşleşen satırı bulduğunda, o satırdaki istediğimiz bilgiyi bize teslim eder.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Sözdizimi ve 4 Hayati Parametre',
      content:
        '=DÜŞEYARA(aranan_değer; tablo_dizisi; sütun_indis_sayısı; [aralık_bak])\n\n1. aranan_değer: Neyi arıyoruz? (örneğin A2\'deki "U103")\n2. tablo_dizisi: Hangi tabloda arıyoruz? (E2:G5)\n3. sütun_indis_sayısı: Kaçıncı sütundaki veriyi istiyoruz? (Katalogda Ürün Adı 2. sütundur)\n4. aralık_bak: Tam eşleşme için 0 (veya YANLIŞ).',
      keyTakeaway: '4 parametre: Neyi ara, Nerede ara, Kaçıncı sütunu getir, 0 (Tam eşleşme).',
      audioNarration:
        'Dört parametre kullanırız: Aradığımız hücre, kaynak tablonun aralığı, getirmek istediğimiz sütun numarası ve tam eşleşme anlamına gelen sıfır rakamı.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Altın Kural: Aranan Değer En Solda Olmalıdır',
      content:
        'DÜŞEYARA yalnızca tablonun 1. sütununda arama yapabilir ve sağa doğru veri getirebilir. Asla sola doğru arama yapamaz. Ayrıca 4. parametre olarak 0 (sıfır) yazılmazsa Excel yaklaşık eşleşme yapar ve yanlış sonuçlar getirebilir.',
      keyTakeaway: 'Aranan anahtar ilk sütunda olmalı ve son parametre daima 0 girilmelidir.',
      audioNarration:
        'DÜŞEYARA kullanırken altın kural şudur: Aradığımız kod tablonun en solundaki ilk sütunda bulunmalıdır ve tam eşleşme için son parametreye daima sıfır yazılmalıdır.',
    },
  ],

  solvedExample: {
    title: 'Ürün Kodu ile Ürün Adını Çekme',
    scenario:
      'A2 hücresinde "U103" kodu yazmaktadır. E2:G5 katalog tablosundan bu kodun karşılığı olan ürün adını B2 hücresine getirmek istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'B2 hücresine tıklayın.',
        tip: 'Sonucun yazılacağı B2 hücresini seçin.',
      },
      {
        stepNumber: 2,
        instruction: '=DÜŞEYARA(A2;E2:G5;2;0) yazıp Enter tuşuna basın.',
        tip: 'İngilizce Excel için =VLOOKUP(A2;E2:G5;2;0) yazabilirsiniz.',
      },
    ],
    explanation:
      'Excel E2:E5 sütununda "U103" kodunu arar, 4. satırda bulur ve tablonun 2. sütunundaki "27 inç Monitör" metnini B2 hücresine yazar.',
    audioNarration:
      'Örneğimizde U103 kodu aranır, katalog tablosundaki ikinci sütundan 27 inç Monitör bilgisi getirilerek ekrana yazılır.',
  },

  realWorldScenarios: [
    {
      industry: 'Perakende & Mağazacılık',
      role: 'Kasa ve Stok Sorumlusu',
      challenge: 'Barkod okutulduğunda 10.000 ürünlük stok listesinden ürün adını ve fiyatını kasada görüntülemek.',
      solution: '=DÜŞEYARA(A2; BarkodTablosu; 3; 0) formülü ile otomatik fiyat dökümü sağlamak.',
      appliedBenefit: 'Kasiyerlerin elle fiyat araması önlendi, satış hızlandı ve kasa kuyrukları ortadan kalktı.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Son parametreyi (0 / YANLIŞ) unutmak',
      cause: 'Son parametre yazılmazsa Excel varsayılan olarak yaklaşık eşleşme (1) yapar ve yanlış satırı getirir.',
      solution: 'İş ve finans tablolarında tam eşleşme için daima 0 yazın: =DÜŞEYARA(...; 0).',
    },
    {
      mistake: 'Sütun indis numarasını sayfa sütun harfi sanmak',
      cause: 'Veri F sütununda diye 6 yazmak yerine tablonun kendi içindeki sütun sırasını (2) saymalısınız.',
      solution: 'Seçtiğiniz tablonun en sol sütununu 1 sayarak sağa doğru numaralandırın.',
    },
  ],

  tasks: [
    {
      id: 'task_vlookup_1',
      number: 1,
      title: '1. Adım: B2 Hücresini Seçin',
      instruction: 'Ürün adını getireceğimiz B2 hücresine tıklayın.',
      narrationText: 'B2 hücresini seçelim ve DÜŞEYARA formülüyle ürün adını çekelim.',
      targetCell: 'B2',
      hints: {
        level1: 'Ürün Adı sütunundaki B2 hücresine tıklayın.',
        level2: 'B sütunu 2. satır.',
        level3: 'B2 hücresine tıklayın.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'B', 2, 'Ürün Adı'),
    },
    {
      id: 'task_vlookup_2',
      number: 2,
      title: '2. Adım: DÜŞEYARA ile Ürün Adını Getirin',
      instruction:
        'B2 hücresine =DÜŞEYARA(A2;E2:G5;2;0) formülünü yazıp Enter tuşuna basın.',
      narrationText:
        'A2 hücresindeki U103 kodunu E2:G5 tablosunda arıyoruz. Ürün adı tablonun 2. sütunundadır.',
      targetCell: 'B2',
      hints: {
        level1: 'Formül kalıbı: =DÜŞEYARA(A2;E2:G5;2;0)',
        level2: 'Tablo aralığı E2:G5, sütun sayısı 2, tam eşleşme için 0.',
        level3: '=DÜŞEYARA(A2;E2:G5;2;0) veya =VLOOKUP(A2;E2:G5;2;0) yazıp Enter\'a basın. Sonuç "27 inç Monitör" olacaktır.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'B2',
          [
            '=DÜŞEYARA(A2;E2:G5;2;0)',
            '=VLOOKUP(A2;E2:G5;2;0)',
            '=DÜŞEYARA(A2,E2:G5,2,0)',
            '=VLOOKUP(A2,E2:G5,2,0)',
          ],
          '27 inç Monitör'
        ),
    },
    {
      id: 'task_vlookup_3',
      number: 3,
      title: '3. Adım: DÜŞEYARA ile Birim Fiyatı Getirin',
      instruction:
        'C2 hücresine tıklayın ve =DÜŞEYARA(A2;E2:G5;3;0) formülünü yazıp Enter tuşuna basın.',
      narrationText:
        'Fiyat tablonun 3. sütununda yer aldığı için bu kez sütun indisi olarak 3 yazıyoruz. Sonuç 4200 ₺ olmalıdır.',
      targetCell: 'C2',
      hints: {
        level1: 'C2 hücresini seçin ve 3. sütundaki fiyatı getiren formülü yazın.',
        level2: '=DÜŞEYARA(A2;E2:G5;3;0) yazın.',
        level3: 'C2 hücresine =DÜŞEYARA(A2;E2:G5;3;0) veya =VLOOKUP(A2;E2:G5;3;0) yazıp Enter\'a basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'C2',
          [
            '=DÜŞEYARA(A2;E2:G5;3;0)',
            '=VLOOKUP(A2;E2:G5;3;0)',
            '=DÜŞEYARA(A2,E2:G5,3,0)',
            '=VLOOKUP(A2,E2:G5,3,0)',
          ],
          4200
        ),
    },
  ],

  independentPractice: {
    goal: 'Personel tablosunda Sicil No "P202" olan personelin departmanını D2:E4 tablosundan B2 hücresine DÜŞEYARA ile çekin.',
    scenario:
      'İK departmanında sicil sorgusu yapıyorsunuz. A2 hücresindeki sicil numarasını D2:E4 tablosunda aratarak 2. sütundaki departman adını bulun.',
    narrationText:
      'Bağımsız alıştırmada P202 sicil numaralı personelin departmanını çekin. Beklenen sonuç "Pazarlama"dır.',
    workbookFactory: createVlookupPracticeWorkbook,
    targetCell: 'B2',
    hints: [
      'B2 hücresini seçin.',
      'Arama tablosu D2:E4, sütun indisi 2, tam eşleşme 0.',
      '=DÜŞEYARA(A2;D2:E4;2;0) yazıp Enter\'a basın.',
    ],
    validate: (sheet: Worksheet) =>
      validateFormulaGeneral(
        sheet,
        'B2',
        [
          '=DÜŞEYARA(A2;D2:E4;2;0)',
          '=VLOOKUP(A2;D2:E4;2;0)',
          '=DÜŞEYARA(A2,D2:E4,2,0)',
          '=VLOOKUP(A2,D2:E4,2,0)',
        ],
        'Pazarlama'
      ),
  },

  quizQuestions: [
    {
      id: 'q_vlookup_1',
      question: 'DÜŞEYARA fonksiyonunda tam eşleşme (birebir aynı değeri arama) sağlamak için son parametre ne olmalıdır?',
      options: ['0 (veya YANLIŞ)', '1 (veya DOĞRU)', '2', 'Boş bırakılmalıdır'],
      correctIndex: 0,
      explanation: 'Tam eşleşme için son parametre daima 0 veya YANLIŞ (FALSE) olarak girilir.',
    },
    {
      id: 'q_vlookup_2',
      question: 'DÜŞEYARA fonksiyonunun İngilizce Excel adı nedir?',
      options: ['VLOOKUP', 'HLOOKUP', 'LOOKUP', 'XLOOKUP'],
      correctIndex: 0,
      explanation: 'DÜŞEYARA fonksiyonunun İngilizce karşılığı VLOOKUP\'tır.',
    },
  ],

  initialWorkbookFactory: createVlookupInitialWorkbook,
};
