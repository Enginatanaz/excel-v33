/**
 * ÇAPRAZARA (XLOOKUP) Fonksiyonu Dersi - Excel Akademi AI
 * Aşama 5: Yeni Nesil Esnek Arama Motoru
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateFormulaGeneral,
} from '../../modules/learning/taskValidator.ts';

export function createXlookupInitialWorkbook(): Workbook {
  return {
    id: 'wb_xlookup_lesson',
    name: 'Yeni_Nesil_Caprazara.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Müşteri Sorgulama',
        rowCount: 15,
        colCount: 8,
        cells: {
          A1: { raw: 'Aranan ID', computed: 'Aranan ID', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          B1: { raw: 'Müşteri Şehri', computed: 'Müşteri Şehri', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          C1: { raw: 'Olmayan ID Testi', computed: 'Olmayan ID Testi', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          A2: { raw: 'C-99', computed: 'C-99', style: { bold: true } },
          B2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },
          C2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },

          // Master Database with Lookup Key on the RIGHT of City (VLOOKUP cannot do this, XLOOKUP can!)
          E1: { raw: 'Şehir (Sol Sütun)', computed: 'Şehir (Sol Sütun)', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          F1: { raw: 'Müşteri ID (Arama)', computed: 'Müşteri ID (Arama)', style: { bold: true, align: 'left', bgColor: '#fef08a' } },
          G1: { raw: 'Bakiye (₺)', computed: 'Bakiye (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          E2: { raw: 'Bursa', computed: 'Bursa' },
          F2: { raw: 'C-10', computed: 'C-10' },
          G2: { raw: '1200', computed: 1200, style: { align: 'right', format: 'currency' } },
          E3: { raw: 'Ankara', computed: 'Ankara' },
          F3: { raw: 'C-55', computed: 'C-55' },
          G3: { raw: '3400', computed: 3400, style: { align: 'right', format: 'currency' } },
          E4: { raw: 'İzmir', computed: 'İzmir' },
          F4: { raw: 'C-99', computed: 'C-99' },
          G4: { raw: '5100', computed: 5100, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export function createXlookupPracticeWorkbook(): Workbook {
  return {
    id: 'wb_xlookup_practice',
    name: 'Bakiye_Sorgulama.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Bakiye Sorgu',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Sorgulanan ID', computed: 'Sorgulanan ID', style: { bold: true, align: 'left', bgColor: '#e0f2fe' } },
          B1: { raw: 'Bulunan Bakiye', computed: 'Bulunan Bakiye', style: { bold: true, align: 'right', bgColor: '#e0f2fe' } },
          A2: { raw: 'C-55', computed: 'C-55', style: { bold: true } },
          B2: { raw: '', computed: null, style: { bgColor: '#f8fafc' } },

          D1: { raw: 'ID', computed: 'ID', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          E1: { raw: 'Hesap Bakiyesi (₺)', computed: 'Hesap Bakiyesi (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          D2: { raw: 'C-10', computed: 'C-10' },
          E2: { raw: '1200', computed: 1200, style: { align: 'right', format: 'currency' } },
          D3: { raw: 'C-55', computed: 'C-55' },
          E3: { raw: '3400', computed: 3400, style: { align: 'right', format: 'currency' } },
          D4: { raw: 'C-99', computed: 'C-99' },
          E4: { raw: '5100', computed: 5100, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export const xlookupFunctionLesson: Lesson = {
  id: 'xlookup_function',
  number: 7,
  title: 'ÇAPRAZARA (XLOOKUP) ile Modern Veri Arama',
  lessonType: 'formula_function',
  category: 'Modern Arama Fonksiyonları',
  level: 'Orta-İleri',
  estimatedMinutes: 16,
  description:
    'DÜŞEYARA\'nın tüm kısıtlamalarını ortadan kaldıran; sola arama yapabilen, varsayılan olarak tam eşleşen ve bulunamadığında özel mesaj veren ÇAPRAZARA fonksiyonunu öğrenin.',
  learningObjective:
    'ÇAPRAZARA (XLOOKUP) fonksiyonunun bağımsız arama ve dönüş aralıklarını kavramak, bulunamıyorsa parametresi ile #YOK! hatalarını profesyonelce yönetebilmek.',
  learningOutcomes: [
    'Arama sütununun solundaki verileri sütun sırası engeli olmadan getirebilme',
    'Varsayılan tam eşleşme avantajını kullanarak 0 parametresi yazma zorunluluğundan kurtulma',
    'Bulunamayan kayıtlar için özel mesaj ("[bulunamıyorsa]") tanımlayabilme',
  ],
  prerequisites: [
    {
      id: 'vlookup_function',
      title: 'DÜŞEYARA (VLOOKUP) Fonksiyonu',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde Excel’in en büyük yeniliklerinden biri olan ÇAPRAZARA (XLOOKUP) fonksiyonunu inceleyeceğiz. DÜŞEYARA\'nın "yalnızca sağa doğru arama" kısıtlamasını yıkan ve hata yönetimini kendi içinde barındıran bu modern fonksiyonu birlikte keşfedelim.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'Neden ÇAPRAZARA (XLOOKUP)?',
      content:
        'DÜŞEYARA 30 yılı aşkın süredir kullanılıyor ancak iki büyük kusuru vardı: Sadece sağa arama yapabilirdi ve sonuna 0 yazmayı unuttuğunuzda yanlış değer getirirdi. ÇAPRAZARA hem sola hem sağa arar, varsayılan olarak tam eşleşme yapar ve bulunamadığında #YOK! yerine "Kayıt Yok" gibi mesajlar gösterebilir.',
      keyTakeaway: 'Sola arayabilir, varsayılan tam eşleşmedir ve hata mesajı verebilir.',
      audioNarration:
        'ÇAPRAZARA fonksiyonu, eski arama fonksiyonlarının tüm eksiklerini kapatır. İster sağa ister sola bakın, güvenli ve tam eşleşmeyle sonuç üretir.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Sözdizimi ve Temel Parametreler',
      content:
        '=ÇAPRAZARA(aranan_değer; arama_aralığı; döndürülen_aralık; [bulunamıyorsa])\n\n• aranan_değer: Arayacağınız anahtar (A2)\n• arama_aralığı: Sadece anahtarın arandığı sütun (F2:F4)\n• döndürülen_aralık: Sonucun çekileceği sütun (E2:E4)\n• bulunamıyorsa: Değer yoksa gösterilecek metin ("Bulunamadı")',
      keyTakeaway: 'Arama aralığı ve dönüş aralığı birbirinden bağımsız iki tek sütun olarak verilir.',
      audioNarration:
        'ÇAPRAZARA formülünde bütün tabloyu seçmek yerine, sadece aranacak sütunu ve sonucu getirecek sütunu ayrı ayrı belirtiriz.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Sola Doğru Arama Üstünlüğü',
      content:
        'Örnek tablomuzda aranan Müşteri ID F sütunundadır, ancak getirmek istediğimiz Şehir bilgisi solundaki E sütunundadır. DÜŞEYARA bu tabloda çalışamazken, ÇAPRAZARA =ÇAPRAZARA(A2;F2:F4;E2:E4) ile doğrudan sol taraftaki veriyi getirebilir.',
      keyTakeaway: 'Sola doğru arama yapabilmesi ÇAPRAZARA\'nın en devrimci özelliğidir.',
      audioNarration:
        'Gelişmiş tablolarda aradığımız kimlik sütununun solunda kalan verileri çekmek için ÇAPRAZARA mükemmel bir çözümdür.',
    },
  ],

  solvedExample: {
    title: 'Müşteri ID ile Sola Arama Yaparak Şehri Bulma',
    scenario:
      'A2 hücresindeki "C-99" kodunu F2:F4 aralığında arayıp sol tarafındaki E2:E4 aralığından şehir adını B2 hücresine çekmek istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'B2 hücresine tıklayın.',
        tip: 'Sonucun yazılacağı B2 hücresini seçin.',
      },
      {
        stepNumber: 2,
        instruction: '=ÇAPRAZARA(A2;F2:F4;E2:E4) formülünü yazıp Enter\'a basın.',
        tip: 'İngilizce Excel için =XLOOKUP(A2;F2:F4;E2:E4) yazabilirsiniz.',
      },
    ],
    explanation:
      'Excel F2:F4 aralığında "C-99" kodunu 4. satırda bulur ve aynı satırın E sütunundaki karşılığı olan "İzmir" metnini getirir.',
    audioNarration:
      'ÇAPRAZARA formülü C-99 kodunu bularak solundaki İzmir şehrini hatasız biçimde ekrana getirir.',
  },

  realWorldScenarios: [
    {
      industry: 'Müşteri Hizmetleri & Çağrı Merkezi',
      role: 'Destek Operatörü',
      challenge: 'Çağrı sırasında müşteri TC kimlik numarasını söylediğinde karmaşık müşteri veri tabanından adı, paketi ve borcunu tek ekranda toplamak.',
      solution: '=ÇAPRAZARA(A2; MusteriTC; MusteriAdi; "Müşteri Kaydı Bulunamadı") formülünü kullanmak.',
      appliedBenefit: 'Yanlış eşleşme riski sıfırlandı, kayıt bulunamadığında çağrı temsilcisine dostça uyarı gösterildi.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Tüm tabloyu tek blok olarak seçmek',
      cause: 'DÜŞEYARA alışkanlığıyla E2:G4 yazıldığında ÇAPRAZARA arama sütunu ile dönüş sütununu ayırt edemez.',
      solution: 'Arama sütununu (F2:F4) ve dönüş sütununu (E2:E4) ayrı ayrı iki bağımsız aralık olarak yazın.',
    },
    {
      mistake: 'Arama ve dönüş aralıklarının boyunu farklı seçmek',
      cause: 'F2:F4 aralığına karşılık E2:E10 seçilirse Excel #DEĞER! üretir.',
      solution: 'Her iki aralığın da aynı satır adedine sahip olduğundan emin olun.',
    },
  ],

  tasks: [
    {
      id: 'task_xlookup_1',
      number: 1,
      title: '1. Adım: B2 Hücresini Seçin',
      instruction: 'Sonucun yazılacağı B2 hücresine tıklayın.',
      narrationText: 'B2 hücresini seçelim.',
      targetCell: 'B2',
      hints: {
        level1: 'Müşteri Şehri sütunundaki B2 hücresine tıklayın.',
        level2: 'B sütunu 2. satır.',
        level3: 'B2 hücresini seçin.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'B', 2, 'Müşteri Şehri'),
    },
    {
      id: 'task_xlookup_2',
      number: 2,
      title: '2. Adım: ÇAPRAZARA ile Sol Sütundaki Şehri Getirin',
      instruction:
        'B2 hücresine =ÇAPRAZARA(A2;F2:F4;E2:E4) yazıp Enter tuşuna basın.',
      narrationText:
        'A2\'deki C-99 kodunu F sütununda arayıp sol tarafındaki E sütunundan İzmir sonucunu getireceğiz.',
      targetCell: 'B2',
      hints: {
        level1: '=ÇAPRAZARA(A2;F2:F4;E2:E4) yazın.',
        level2: 'Aranan: A2, arama alanı: F2:F4, dönüş alanı: E2:E4.',
        level3: 'B2 hücresine =ÇAPRAZARA(A2;F2:F4;E2:E4) veya =XLOOKUP(A2;F2:F4;E2:E4) yazıp Enter\'a basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'B2',
          [
            '=ÇAPRAZARA(A2;F2:F4;E2:E4)',
            '=XLOOKUP(A2;F2:F4;E2:E4)',
            '=ÇAPRAZARA(A2,F2:F4,E2:E4)',
            '=XLOOKUP(A2,F2:F4,E2:E4)',
          ],
          'İzmir'
        ),
    },
    {
      id: 'task_xlookup_3',
      number: 3,
      title: '3. Adım: Bulunamıyorsa Parametresi ile Güvenli Arama',
      instruction:
        'C2 hücresine tıklayın ve =ÇAPRAZARA("C-999";F2:F4;E2:E4;"Kayıt Yok") yazıp Enter\'a basın.',
      narrationText:
        'Listede olmayan "C-999" arandığında ÇAPRAZARA\'nın #YOK! yerine "Kayıt Yok" döndürdüğünü gözlemleyin.',
      targetCell: 'C2',
      hints: {
        level1: 'C2 hücresini seçin ve 4. parametre olarak "Kayıt Yok" ekleyin.',
        level2: '=ÇAPRAZARA("C-999";F2:F4;E2:E4;"Kayıt Yok") yazın.',
        level3: 'C2 hücresine =ÇAPRAZARA("C-999";F2:F4;E2:E4;"Kayıt Yok") veya =XLOOKUP("C-999";F2:F4;E2:E4;"Kayıt Yok") yazıp Enter\'a basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'C2',
          [
            '=ÇAPRAZARA("C-999";F2:F4;E2:E4;"Kayıt Yok")',
            '=XLOOKUP("C-999";F2:F4;E2:E4;"Kayıt Yok")',
            '=ÇAPRAZARA("C-999",F2:F4,E2:E4,"Kayıt Yok")',
            '=XLOOKUP("C-999",F2:F4,E2:E4,"Kayıt Yok")',
          ],
          'Kayıt Yok'
        ),
    },
  ],

  independentPractice: {
    goal: 'Bakiye Sorgulama tablosunda A2 hücresindeki "C-55" kodunu D2:D4 aralığında aratarak E2:E4 aralığındaki bakiyeyi B2 hücresine çekin.',
    scenario:
      'Müşterinin hesap bakiyesini modern ÇAPRAZARA fonksiyonu ile sorgulayın. Beklenen bakiye 3400 ₺\'dir.',
    narrationText:
      'Bağımsız alıştırmada C-55 kodunun bakiyesini ÇAPRAZARA ile getirin.',
    workbookFactory: createXlookupPracticeWorkbook,
    targetCell: 'B2',
    hints: [
      'B2 hücresini seçin.',
      'Aranan: A2, Arama aralığı: D2:D4, Dönüş aralığı: E2:E4.',
      '=ÇAPRAZARA(A2;D2:D4;E2:E4) yazıp Enter\'a basın.',
    ],
    validate: (sheet: Worksheet) =>
      validateFormulaGeneral(
        sheet,
        'B2',
        [
          '=ÇAPRAZARA(A2;D2:D4;E2:E4)',
          '=XLOOKUP(A2;D2:D4;E2:E4)',
          '=ÇAPRAZARA(A2,D2:D4,E2:E4)',
          '=XLOOKUP(A2,D2:D4,E2:E4)',
        ],
        3400
      ),
  },

  quizQuestions: [
    {
      id: 'q_xlookup_1',
      question: 'ÇAPRAZARA fonksiyonunun DÜŞEYARA\'ya göre en belirgin üstünlüklerinden biri nedir?',
      options: [
        'Arama sütununun solunda kalan sütunlardaki verileri de çekebilmesi',
        'Yalnızca sayıları arayabilmesi',
        'Toplama işlemi yapabilmesi',
        'Sayfa isimlerini değiştirebilmesi',
      ],
      correctIndex: 0,
      explanation: 'ÇAPRAZARA, arama sütununun solunda veya sağında bulunan herhangi bir sütundaki veriyi rahatlıkla getirebilir.',
    },
    {
      id: 'q_xlookup_2',
      question: 'ÇAPRAZARA fonksiyonunda aranan değer bulunamadığında varsayılan hata yerine özel mesaj göstermek için hangi parametre kullanılır?',
      options: [
        '4. parametre ([bulunamıyorsa])',
        '1. parametre',
        'Böyle bir parametre yoktur',
        'Son parametreye daima 0 yazılmalıdır',
      ],
      correctIndex: 0,
      explanation: 'ÇAPRAZARA\'nın 4. parametresi [bulunamıyorsa] (if_not_found), eşleşme olmadığında gösterilecek özel mesajı belirler.',
    },
  ],

  initialWorkbookFactory: createXlookupInitialWorkbook,
};
