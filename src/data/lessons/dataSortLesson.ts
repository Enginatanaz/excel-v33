/**
 * Veri Sıralama Dersi - Excel Akademi AI
 * Aşama 5: Veri Yönetimi & Satır Bütünlüğü
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateTableSorted,
} from '../../modules/learning/taskValidator.ts';

export function createSortInitialWorkbook(): Workbook {
  return {
    id: 'wb_sort_lesson',
    name: 'Personel_Maas_Listesi.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Personel Listesi',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Personel Adı', computed: 'Personel Adı', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Departman', computed: 'Departman', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Maaş (₺)', computed: 'Maaş (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Zeynep Kaya', computed: 'Zeynep Kaya' },
          B2: { raw: 'Pazarlama', computed: 'Pazarlama' },
          C2: { raw: '32000', computed: 32000, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Ahmet Yılmaz', computed: 'Ahmet Yılmaz' },
          B3: { raw: 'Finans', computed: 'Finans' },
          C3: { raw: '45000', computed: 45000, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Canan Şen', computed: 'Canan Şen' },
          B4: { raw: 'İnsan Kaynakları', computed: 'İnsan Kaynakları' },
          C4: { raw: '28000', computed: 28000, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'Burak Aras', computed: 'Burak Aras' },
          B5: { raw: 'Yazılım', computed: 'Yazılım' },
          C5: { raw: '52000', computed: 52000, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export function createSortPracticeWorkbook(): Workbook {
  return {
    id: 'wb_sort_practice',
    name: 'Sehir_Nufus_Verileri.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Şehir Nüfusları',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Şehir', computed: 'Şehir', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Bölge', computed: 'Bölge', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Nüfus', computed: 'Nüfus', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'İzmir', computed: 'İzmir' },
          B2: { raw: 'Ege', computed: 'Ege' },
          C2: { raw: '4400000', computed: 4400000, style: { align: 'right' } },
          A3: { raw: 'Ankara', computed: 'Ankara' },
          B3: { raw: 'İç Anadolu', computed: 'İç Anadolu' },
          C3: { raw: '5800000', computed: 5800000, style: { align: 'right' } },
          A4: { raw: 'Bursa', computed: 'Bursa' },
          B4: { raw: 'Marmara', computed: 'Marmara' },
          C4: { raw: '3100000', computed: 3100000, style: { align: 'right' } },
          A5: { raw: 'İstanbul', computed: 'İstanbul' },
          B5: { raw: 'Marmara', computed: 'Marmara' },
          C5: { raw: '15900000', computed: 15900000, style: { align: 'right' } },
        },
      },
    ],
  };
}

export const dataSortLesson: Lesson = {
  id: 'data_sort_lesson',
  number: 8,
  title: 'Veri Sıralama (A-Z ve Z-A) ve Satır Bütünlüğü',
  lessonType: 'data_management',
  category: 'Veri Yönetimi & Düzenleme',
  level: 'Başlangıç-Orta',
  estimatedMinutes: 14,
  description:
    'Tablolarınızdaki satırları isim sırasına (A-Z) veya sayısal büyüklüğe (Z-A) göre satır bütünlüğünü asla bozmadan profesyonelce sıralamayı öğrenin.',
  learningObjective:
    'Sıralama yaparken tablonun tamamının satır satır hareket ettiğini (satır bütünlüğü), başlık satırının sabit kalması gerektiğini ve metin/sayı sıralama mantığını kavramak.',
  learningOutcomes: [
    'A\'dan Z\'ye alfabetik sıralama yapabilme',
    'Z\'den A\'ya büyükten küçüğe sayısal sıralama yapabilme',
    'Satır bütünlüğünü koruyarak personelin maaş veya departman bilgisiyle ayrışmasını önleme',
  ],
  prerequisites: [
    {
      id: 'excel_intro',
      title: "Excel'e Giriş & Hücre Mantığı",
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde büyük veri tablolarını anlamlandırmanın en temel yöntemi olan Veri Sıralama konusunu ele alacağız. Personel listesini alfabetik sıraya dizmeyi ve en yüksek maaştan en düşüğe doğru sıralamayı öğreneceğiz.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'Veri Sıralamanın Önemi ve Satır Bütünlüğü',
      content:
        'Sıralama, karışık verileri anlamlı bir düzene sokar. Ancak Excel\'de sıralamanın en kritik kuralı SATIR BÜTÜNLÜĞÜDÜR. Bir personelin adı değiştiğinde yanındaki maaşı ve departmanı da onunla birlikte hareket etmelidir. Yalnızca tek bir sütunu seçip sıralamak tüm tabloyu bozar!',
      keyTakeaway: 'Sıralama yapıldığında satırdaki tüm ilişkili veriler birlikte hareket etmelidir.',
      audioNarration:
        'Sıralama yaparken en önemli prensip satır bütünlüğüdür. Bir kişinin adını sıraladığınızda, o kişinin maaşı ve departmanı da onunla beraber aynı satırda kalmalıdır.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'A-Z ve Z-A Sıralama Mantığı',
      content:
        '• A-Z: Metin sütunlarında alfabetik (A\'dan Z\'ye), sayısal sütunlarda küçükten büyüğe (1\'den 100\'e) sıralar.\n• Z-A: Metin sütunlarında ters alfabetik (Z\'den A\'ya), sayısal sütunlarda büyükten küçüğe (en yüksek cirodan en düşüğe) sıralar.',
      keyTakeaway: 'A-Z küçükten büyüğe/alfabetik, Z-A büyükten küçüğe sıralar.',
      audioNarration:
        'A-Z sıralaması isimleri alfabetik, sayıları ise küçükten büyüğe dizer. Z-A ise tam tersine en yüksek sayıları veya ters alfabetik sırayı en başa getirir.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Başlık Satırının Korunması',
      content:
        'Tablonun ilk satırındaki başlıklar ("Personel Adı", "Maaş") birer veri değildir. Excel akıllı algoritmaları bu satırı tespit eder ve başlıkların aşağı kaymasını engelleyerek sıralamayı yalnızca veri satırları (2-5) arasında gerçekleştirir.',
      keyTakeaway: 'Başlık satırı sabit kalır, veriler kendi aralarında sıralanır.',
      audioNarration:
        'Tablo başlıkları veri olarak kabul edilmez. Sıralama motorumuz başlıkların sabit kalmasını garanti eder.',
    },
  ],

  solvedExample: {
    title: 'Personel Listesini İsim Sırasına Koyma',
    scenario:
      'Karışık olan personel listesini A sütununa göre A-Z alfabetik sıraya dizmek istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'A sütunundaki herhangi bir personel hücresine (örneğin A2) tıklayın.',
        tip: 'Sıralamak istediğiniz sütunda bir hücre seçin.',
      },
      {
        stepNumber: 2,
        instruction: 'Üst araç çubuğundaki "A-Z Sırala" butonuna tıklayın.',
        tip: 'Veya sütun başlığındaki açılır menüden A\'dan Z\'ye Sırala seçeneğini kullanın.',
      },
    ],
    explanation:
      'Ahmet Yılmaz en üste, Burak Aras ikinci sıraya, Canan Şen üçüncü sıraya ve Zeynep Kaya dördüncü sıraya geçer. Departman ve maaşları da onlarla birlikte taşınır.',
    audioNarration:
      'A2 hücresini seçip A-Z sırala butonuna bastığımızda tüm personel isim sırasına geçer ve satır bütünlüğü kusursuz biçimde korunur.',
  },

  realWorldScenarios: [
    {
      industry: 'İnsan Kaynakları & Bordro',
      role: 'Bordro Uzmanı',
      challenge: '300 çalışanın maaş listesini en yüksekten en düşüğe sıralayarak bütçe dağılımını analiz etmek.',
      solution: 'Maaş sütununda Z-A (Büyükten Küçüğe) sıralama yaparak üst yönetim raporunu hazırlamak.',
      appliedBenefit: 'Maaşlar satırlarından kopmadan hatasız olarak en yüksekten en düşüğe dizildi.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Sıralama öncesinde yalnızca tek sütundaki hücreleri seçmek',
      cause: 'Eski Excel sürümlerinde tek sütun seçildiğinde diğer sütunlar sabit kalır ve Ahmet\'in maaşı Zeynep\'e geçer!',
      solution: 'Tek bir hücre seçili bırakın veya tablonun tamamını kapsayacak şekilde sıralayın.',
    },
    {
      mistake: 'Başlık satırını da veriye dahil ederek sıralamak',
      cause: 'Başlık algılanmazsa "Personel Adı" metni P harfi hizasına kayabilir.',
      solution: 'Başlık satırını kalın (bold) veya renkli biçimlendirerek Excel\'in başlığı tanımasını sağlayın.',
    },
  ],

  tasks: [
    {
      id: 'task_sort_1',
      number: 1,
      title: '1. Adım: A2 Hücresini Seçin',
      instruction: 'Personel Adı sütunundaki A2 hücresine tıklayın.',
      narrationText: 'A2 hücresini seçelim.',
      targetCell: 'A2',
      hints: {
        level1: 'A2 hücresine tıklayın.',
        level2: 'Personel Adı sütununun ilk satırı.',
        level3: 'A2 hücresini seçin.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'A', 2, 'Personel Adı'),
    },
    {
      id: 'task_sort_2',
      number: 2,
      title: '2. Adım: Tabloyu A\'dan Z\'ye Sıralayın',
      instruction:
        'Üst araç çubuğundaki "A-Z Sırala" butonuna (veya A sütun başlığındaki açılır menüye) tıklayarak isimleri alfabetik sıralayın.',
      narrationText:
        'A-Z butonuna tıklayarak isimleri alfabetik dizin. Ahmet Yılmaz en başa gelecektir.',
      targetCell: 'A2',
      hints: {
        level1: 'Araç çubuğundaki A-Z simgeli Sırala butonuna tıklayın.',
        level2: 'Aktif hücreniz A sütunundayken A-Z Sırala butonuna basın.',
        level3: 'A sütunu için A-Z butonuna tıklayın.',
      },
      validate: (sheet) => validateTableSorted(sheet, 'A', true, { startRow: 2, endRow: 5 }),
    },
    {
      id: 'task_sort_3',
      number: 3,
      title: '3. Adım: Maaş Sütununu Büyükten Küçüğe (Z-A) Sıralayın',
      instruction:
        'C2 hücresine tıklayın ve araç çubuğundaki "Z-A Sırala" butonuna tıklayarak en yüksek maaşı en üste getirin.',
      narrationText:
        'Maaş sütununa geçin ve Z-A butonuna basarak büyükten küçüğe sıralayın. 52.000 ₺ ile Burak Aras en üste geçmelidir.',
      targetCell: 'C2',
      hints: {
        level1: 'C sütununda bir hücre seçin ve Z-A butonuna tıklayın.',
        level2: 'C2 hücresine tıklayın, ardından araç çubuğundaki Z-A Sırala butonuna basın.',
        level3: 'C sütununu Z-A sıralayın.',
      },
      validate: (sheet) => validateTableSorted(sheet, 'C', false, { startRow: 2, endRow: 5 }),
    },
  ],

  independentPractice: {
    goal: 'Şehir Nüfusları tablosunda C sütunundaki Nüfus verilerini en kalabalıktan en aza doğru (Z-A / Büyükten Küçüğe) sıralayın.',
    scenario:
      'Nüfus verilerini analiz eden bir şehir plancısısınız. C2:C5 alanını Z-A sıralayarak İstanbul\'u en üste getirin.',
    narrationText:
      'Bağımsız alıştırmada C sütununu Z-A sıralayarak en büyük nüfusu tepeye taşıyın.',
    workbookFactory: createSortPracticeWorkbook,
    targetCell: 'C2',
    hints: [
      'C sütunundaki herhangi bir hücreye (örneğin C2) tıklayın.',
      'Araç çubuğundaki Z-A Sırala butonuna tıklayın.',
      'Sıralama sonucunda en üstte 15.900.000 ile İstanbul yer almalıdır.',
    ],
    validate: (sheet) => validateTableSorted(sheet, 'C', false, { startRow: 2, endRow: 5 }),
  },

  quizQuestions: [
    {
      id: 'q_sort_1',
      question: 'Excel\'de bir sütuna göre sıralama yapıldığında diğer sütunlardaki veriler ne olur?',
      options: [
        'Satır bütünlüğü korunarak ilgili satırla birlikte hareket eder',
        'Oldukları yerde sabit kalırlar ve satırlar karışır',
        'Tüm diğer sütunlar silinir',
        'Diğer sütunlar otomatik olarak alfabetik sıralanır',
      ],
      correctIndex: 0,
      explanation: 'Excel doğru tasarlandığında satır bütünlüğünü korur; her satırdaki hücreler bir bütün olarak taşınır.',
    },
    {
      id: 'q_sort_2',
      question: 'Sayısal bir sütunda Z-A sıralaması yapıldığında sonuç nasıl sıralanır?',
      options: [
        'En büyük sayıdan en küçük sayıya (Büyükten Küçüğe)',
        'En küçük sayıdan en büyük sayıya (Küçükten Büyüğe)',
        'Sayıların basamak sayısına göre',
        'Rastgele sırayla',
      ],
      correctIndex: 0,
      explanation: 'Sayısal verilerde Z-A seçeneği büyükten küçüğe doğru (azalan) sıralama yapar.',
    },
  ],

  initialWorkbookFactory: createSortInitialWorkbook,
};
