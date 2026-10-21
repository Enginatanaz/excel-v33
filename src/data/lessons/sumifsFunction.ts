/**
 * ÇOKETOPLA (SUMIFS) Fonksiyonu Dersi - Excel Akademi AI
 * Aşama 5: Çoklu Ölçütlü Koşullu Toplama
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateFormulaGeneral,
} from '../../modules/learning/taskValidator.ts';

export function createSumifsInitialWorkbook(): Workbook {
  return {
    id: 'wb_sumifs_lesson',
    name: 'Coklu_Kriter_Satis_Analizi.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Detaylı Satışlar',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Temsilci', computed: 'Temsilci', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Şehir', computed: 'Şehir', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Kategori', computed: 'Kategori', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          D1: { raw: 'Tutar (₺)', computed: 'Tutar (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Can', computed: 'Can' },
          B2: { raw: 'İstanbul', computed: 'İstanbul' },
          C2: { raw: 'Elektronik', computed: 'Elektronik' },
          D2: { raw: '5000', computed: 5000, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Ece', computed: 'Ece' },
          B3: { raw: 'Ankara', computed: 'Ankara' },
          C3: { raw: 'Gıda', computed: 'Gıda' },
          D3: { raw: '2000', computed: 2000, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Can', computed: 'Can' },
          B4: { raw: 'İstanbul', computed: 'İstanbul' },
          C4: { raw: 'Gıda', computed: 'Gıda' },
          D4: { raw: '3000', computed: 3000, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'Murat', computed: 'Murat' },
          B5: { raw: 'İzmir', computed: 'İzmir' },
          C5: { raw: 'Elektronik', computed: 'Elektronik' },
          D5: { raw: '4500', computed: 4500, style: { align: 'right', format: 'currency' } },
          A6: { raw: 'Can', computed: 'Can' },
          B6: { raw: 'İstanbul', computed: 'İstanbul' },
          C6: { raw: 'Elektronik', computed: 'Elektronik' },
          D6: { raw: '7000', computed: 7000, style: { align: 'right', format: 'currency' } },
          A8: { raw: 'Can - Elektronik Toplamı:', computed: 'Can - Elektronik Toplamı:', style: { bold: true } },
          B8: { raw: '', computed: null, style: { bold: true, align: 'right', bgColor: '#e2e8f0' } },
        },
      },
    ],
  };
}

export function createSumifsPracticeWorkbook(): Workbook {
  return {
    id: 'wb_sumifs_practice',
    name: 'Bolge_Hedef_Toplami.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Bölge Satışları',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Bölge', computed: 'Bölge', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Kanal', computed: 'Kanal', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Ciro (₺)', computed: 'Ciro (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Marmara', computed: 'Marmara' },
          B2: { raw: 'Online', computed: 'Online' },
          C2: { raw: '8000', computed: 8000, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Ege', computed: 'Ege' },
          B3: { raw: 'Mağaza', computed: 'Mağaza' },
          C3: { raw: '6000', computed: 6000, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Marmara', computed: 'Marmara' },
          B4: { raw: 'Mağaza', computed: 'Mağaza' },
          C4: { raw: '4000', computed: 4000, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'Marmara', computed: 'Marmara' },
          B5: { raw: 'Online', computed: 'Online' },
          C5: { raw: '12000', computed: 12000, style: { align: 'right', format: 'currency' } },
          A7: { raw: 'Marmara Online Toplamı:', computed: 'Marmara Online Toplamı:', style: { bold: true } },
          B7: { raw: '', computed: null, style: { bold: true, align: 'right', bgColor: '#e2e8f0' } },
        },
      },
    ],
  };
}

export const sumifsFunctionLesson: Lesson = {
  id: 'sumifs_function',
  number: 5,
  title: 'ÇOKETOPLA (SUMIFS) ile Çok Kriterli Toplama',
  lessonType: 'formula_function',
  category: 'Gelişmiş Formüller',
  level: 'Orta',
  estimatedMinutes: 16,
  description:
    'Birden fazla şarta aynı anda uyan satırların sayısal toplamını alan ÇOKETOPLA (SUMIFS) fonksiyonunu uzman seviyesinde öğrenin.',
  learningObjective:
    'ÇOKETOPLA fonksiyonunda toplama aralığının en başa yazıldığını kavramak, çoklu ölçüt aralığı ve ölçüt çiftlerini doğru sırayla birleştirebilmek.',
  learningOutcomes: [
    'ETOPLA ile ÇOKETOPLA arasındaki parametre sırası farkını anlama',
    'Birden fazla bağımsız filtreleme koşulunu tek bir formülde toplayabilme',
    'ÇOKETOPLA ve SUMIFS formüllerini finansal ve operasyonel raporlarda uygulayabilme',
  ],
  prerequisites: [
    {
      id: 'sum_function',
      title: 'TOPLA (SUM) Fonksiyonu',
    },
    {
      id: 'countif_function',
      title: 'EĞERSAY (COUNTIF) Fonksiyonu',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde iş hayatında finans ve satış analistlerinin en çok kullandığı fonksiyonlardan biri olan ÇOKETOPLA fonksiyonunu inceleyeceğiz. Hem belirli bir temsilcinin hem de belirli bir ürün kategorisinin satışlarını aynı anda toplayarak çok boyutlu raporlar hazırlayacağız.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'ÇOKETOPLA Fonksiyonunun Amacı',
      content:
        'ÇOKETOPLA, aynı anda birden fazla koşula uyan hücreleri toplar. Örneğin "İstanbul\'da yapılan VE Elektronik kategorisindeki VE Temsilcisi Can olan satışların toplamı" gibi karmaşık sorguları tek bir formülle çözer.',
      keyTakeaway: 'Birden fazla koşulun tümünü (VE mantığıyla) sağlayan satırları toplar.',
      audioNarration:
        'ÇOKETOPLA fonksiyonu, belirttiğimiz şartların hepsini aynı anda karşılayan satırları bulur ve sayısal toplamlarını çıkarır.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Kritik Sözdizimi Kuralı: Toplam Aralığı En Baştadır',
      content:
        '=ÇOKETOPLA(toplam_aralığı; ölçüt_aralığı1; ölçüt1; [ölçüt_aralığı2; ölçüt2]; ...)\n\nDikkat: Tek ölçütlü ETOPLA fonksiyonunda toplanacak alan en sondadır, ancak ÇOKETOPLA fonksiyonunda toplanacak sayı sütunu İLK parametre olarak yazılır!',
      keyTakeaway: 'En büyük tuzak: Toplanacak sayılar ilk parametredir.',
      audioNarration:
        'ÇOKETOPLA fonksiyonunda en dikkat edilmesi gereken nokta şudur: Toplanacak sayıların bulunduğu aralık en başa yazılır. Ardından kriter alanları ve kriter değerleri gelir.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Aralık Boyutlarının Birebir Eşitliği',
      content:
        'Toplam aralığı D2:D6 ise, tüm ölçüt aralıkları da mutlaka 2. satırdan başlayıp 6. satırda bitmelidir (örneğin A2:A6 ve C2:C6). Aralık boyutları farklı olursa Excel #DEĞER! hatası verir.',
      keyTakeaway: 'Tüm ölçüt ve toplam aralıkları aynı satır sayısına sahip olmalıdır.',
      audioNarration:
        'Kullandığımız bütün aralıkların satır sayıları eşit olmak zorundadır. Aksi halde Excel hesaplama hatası üretir.',
    },
  ],

  solvedExample: {
    title: 'Can\'ın Elektronik Satışları Toplamı',
    scenario:
      'Temsilcisi "Can" VE kategorisi "Elektronik" olan satış tutarlarının toplamını B8 hücresinde hesaplamak istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'B8 hücresine tıklayın.',
        tip: 'Sonuç hücresini seçin.',
      },
      {
        stepNumber: 2,
        instruction:
          '=ÇOKETOPLA(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") formülünü yazıp Enter\'a basın.',
        tip: 'İngilizce Excel için =SUMIFS(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") kullanın.',
      },
    ],
    explanation:
      'Can\'ın Elektronik kategorisindeki satışları 2. satırda 5000 ₺ ve 6. satırda 7000 ₺\'dir. Formül bu iki sayıyı toplar ve 12000 ₺ sonucunu verir.',
    audioNarration:
      'Örneğimizde Can temsilcisine ait Elektronik satışları toplanır: beş bin artı yedi bin, sonuç tam on iki bin liradır.',
  },

  realWorldScenarios: [
    {
      industry: 'Finans & Muhasebe',
      role: 'Finansal Kontrolör',
      challenge: 'Gider dökümünde "Pazarlama" departmanına ait VE "Yazılım" harcaması olan fatura toplamlarını raporlamak.',
      solution: '=ÇOKETOPLA(E2:E500; B2:B500; "Pazarlama"; C2:C500; "Yazılım") formülü ile dinamik bütçe tablosu kurmak.',
      appliedBenefit: 'Ay sonu kapanış raporları manuel filtrelemeye gerek kalmadan otomatik olarak tamamlandı.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Toplam aralığını en sona yazmak',
      cause: 'ETOPLA alışkanlığıyla toplanacak alanı en sona koymak #DEĞER! veya yanlış hesaplamaya sebep olur.',
      solution: 'ÇOKETOPLA\'da her zaman önce toplanacak sayı sütununu girin: =ÇOKETOPLA(D2:D6; ...).',
    },
    {
      mistake: 'Farklı uzunlukta aralıklar seçmek',
      cause: 'D2:D6 ile A2:A10 eşleştirilmeye çalışıldığında Excel satırları birebir eşleştiremez.',
      solution: 'Tüm aralıkların satır sınırlarının aynı olduğundan (2 ile 6) emin olun.',
    },
  ],

  tasks: [
    {
      id: 'task_sumifs_1',
      number: 1,
      title: '1. Adım: B8 Hücresini Seçin',
      instruction: 'Sonucun yazılacağı B8 hücresine tıklayın.',
      narrationText: 'B8 hücresini seçip formül girişini hazırlayalım.',
      targetCell: 'B8',
      hints: {
        level1: 'Can - Elektronik Toplamı başlığının yanındaki boş hücreye tıklayın.',
        level2: 'B sütunu 8. satırdaki B8 hücresine tıklayın.',
        level3: 'B8 hücresini seçin.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'B', 8, 'Toplam Sonuç Hücresi'),
    },
    {
      id: 'task_sumifs_2',
      number: 2,
      title: '2. Adım: ÇOKETOPLA Formülünü Uygulayın',
      instruction:
        'B8 hücresine =ÇOKETOPLA(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") formülünü yazıp Enter tuşuna basın.',
      narrationText:
        'İlk parametre olarak tutar aralığını (D2:D6), ardından temsilci ölçütünü (A2:A6;"Can") ve kategori ölçütünü (C2:C6;"Elektronik") belirterek toplamı hesaplayın.',
      targetCell: 'B8',
      hints: {
        level1: 'İlk parametre D2:D6 (tutar), ikinci A2:A6 ("Can"), üçüncü C2:C6 ("Elektronik").',
        level2: '=ÇOKETOPLA(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") yazın.',
        level3: 'B8 hücresine =ÇOKETOPLA(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") veya =SUMIFS(D2:D6;A2:A6;"Can";C2:C6;"Elektronik") yazıp Enter\'a basın. Sonuç 12000 olacaktır.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'B8',
          [
            '=ÇOKETOPLA(D2:D6;A2:A6;"Can";C2:C6;"Elektronik")',
            '=SUMIFS(D2:D6;A2:A6;"Can";C2:C6;"Elektronik")',
            '=ÇOKETOPLA(D2:D6,A2:A6,"Can",C2:C6,"Elektronik")',
            '=SUMIFS(D2:D6,A2:A6,"Can",C2:C6,"Elektronik")',
          ],
          12000
        ),
    },
  ],

  independentPractice: {
    goal: 'Bölge Satışları tablosunda Bölgesi "Marmara" VE Kanalı "Online" olan satışların ciro toplamını B7 hücresinde hesaplayın.',
    scenario:
      'Marmara bölgesindeki Online satış hacmini görmek istiyorsunuz. ÇOKETOPLA ile Ciro sütununu (C2:C5), Bölge (A2:A5) ve Kanal (B2:B5) kriterlerine göre toplayın.',
    narrationText:
      'Bağımsız alıştırmada Marmara ve Online koşulunu sağlayan ciroları toplayın. Beklenen sonuç yirmi bindir (8000 + 12000).',
    workbookFactory: createSumifsPracticeWorkbook,
    targetCell: 'B7',
    hints: [
      'B7 hücresini seçin.',
      'Toplam aralığı: C2:C5, Kriter 1: A2:A5;"Marmara", Kriter 2: B2:B5;"Online".',
      '=ÇOKETOPLA(C2:C5;A2:A5;"Marmara";B2:B5;"Online") yazıp Enter\'a basın.',
    ],
    validate: (sheet) =>
      validateFormulaGeneral(
        sheet,
        'B7',
        [
          '=ÇOKETOPLA(C2:C5;A2:A5;"Marmara";B2:B5;"Online")',
          '=SUMIFS(C2:C5;A2:A5;"Marmara";B2:B5;"Online")',
          '=ÇOKETOPLA(C2:C5,A2:A5,"Marmara",B2:B5,"Online")',
          '=SUMIFS(C2:C5,A2:A5,"Marmara",B2:B5,"Online")',
        ],
        20000
      ),
  },

  quizQuestions: [
    {
      id: 'q_sumifs_1',
      question: 'ÇOKETOPLA fonksiyonunda toplanacak sayısal hücrelerin bulunduğu aralık sözdiziminde nerede yer alır?',
      options: [
        'En başta (ilk parametre olarak)',
        'En sonda (son parametre olarak)',
        'Ortada, kriterlerden hemen önce',
        'Fark etmez, Excel kendisi anlar',
      ],
      correctIndex: 0,
      explanation: 'ÇOKETOPLA fonksiyonunda toplanacak aralık (toplam_aralığı) daima ilk parametre olarak yazılır.',
    },
    {
      id: 'q_sumifs_2',
      question: 'ÇOKETOPLA fonksiyonunun İngilizce karşılığı nedir?',
      options: ['SUMIFS', 'SUMIF', 'MULTISUM', 'COUNTIFS'],
      correctIndex: 0,
      explanation: 'ÇOKETOPLA fonksiyonunun İngilizce adı SUMIFS\'tir.',
    },
  ],

  initialWorkbookFactory: createSumifsInitialWorkbook,
};
