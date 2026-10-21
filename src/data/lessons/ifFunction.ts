/**
 * EĞER (IF) Fonksiyonu Dersi - Excel Akademi AI
 * Aşama 5: Formül Motoru & Mantıksal Koşullandırma
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateCellSelection,
  validateFormulaGeneral,
} from '../../modules/learning/taskValidator.ts';

export function createIfInitialWorkbook(): Workbook {
  return {
    id: 'wb_if_lesson',
    name: 'Ogrenci_Not_Degerlendirme.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Not Listesi',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Öğrenci Adı', computed: 'Öğrenci Adı', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Sınav Notu', computed: 'Sınav Notu', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          C1: { raw: 'Durum', computed: 'Durum', style: { bold: true, align: 'center', bgColor: '#f1f5f9' } },
          A2: { raw: 'Ahmet Yılmaz', computed: 'Ahmet Yılmaz' },
          B2: { raw: '75', computed: 75, style: { align: 'right' } },
          C2: { raw: '', computed: null },
          A3: { raw: 'Zeynep Kaya', computed: 'Zeynep Kaya' },
          B3: { raw: '42', computed: 42, style: { align: 'right' } },
          C3: { raw: '', computed: null },
          A4: { raw: 'Mehmet Demir', computed: 'Mehmet Demir' },
          B4: { raw: '88', computed: 88, style: { align: 'right' } },
          C4: { raw: '', computed: null },
          A5: { raw: 'Elif Şahin', computed: 'Elif Şahin' },
          B5: { raw: '50', computed: 50, style: { align: 'right' } },
          C5: { raw: '', computed: null },
        },
      },
    ],
  };
}

export function createIfPracticeWorkbook(): Workbook {
  return {
    id: 'wb_if_practice',
    name: 'Prim_Hak_Edisi.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Satış Hedefleri',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Temsilci', computed: 'Temsilci', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Aylık Satış (₺)', computed: 'Aylık Satış (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          C1: { raw: 'Prim Durumu', computed: 'Prim Durumu', style: { bold: true, align: 'center', bgColor: '#f1f5f9' } },
          A2: { raw: 'Burak Can', computed: 'Burak Can' },
          B2: { raw: '15000', computed: 15000, style: { align: 'right', format: 'currency' } },
          C2: { raw: '', computed: null },
          A3: { raw: 'Selin Aras', computed: 'Selin Aras' },
          B3: { raw: '8500', computed: 8500, style: { align: 'right', format: 'currency' } },
          C3: { raw: '', computed: null },
        },
      },
    ],
  };
}

export const ifFunctionLesson: Lesson = {
  id: 'if_function',
  number: 3,
  title: 'EĞER (IF) Fonksiyonu ile Mantıksal Kararlar',
  lessonType: 'formula_function',
  category: 'Mantıksal Fonksiyonlar',
  level: 'Temel-Orta',
  estimatedMinutes: 14,
  description:
    'Excel’de belirli bir şartın doğru ya da yanlış olmasına göre farklı sonuçlar döndüren EĞER (IF) fonksiyonunu öğrenin.',
  learningObjective:
    'EĞER fonksiyonunun 3 temel parametresini (mantıksal test, doğruysa değer, yanlışsa değer) kavramak ve iş tablolarında koşullu karar mekanizmaları kurabilmek.',
  learningOutcomes: [
    'Mantıksal karşılaştırma işleçlerini (>=, <=, >, <, =, <>) hatasız kullanabilme',
    'EĞER ve İngilizce karşılığı IF sözdizimini metin veya sayısal sonuçlarla uygulayabilme',
    'Koşul sağlandığında ve sağlanmadığında farklı sonuçlar üretebilme',
  ],
  prerequisites: [
    {
      id: 'sum_function',
      title: 'TOPLA (SUM) Fonksiyonu',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde Excel’in en güçlü mantıksal aracı olan EĞER fonksiyonunu keşfedeceğiz. EĞER fonksiyonu sayesinde tablolarınızdaki sayılara bakarak otomatik kararlar verebilirsiniz: Not 50’den büyükse "Geçti", değilse "Kaldı" gibi dinamik kuralları kolayca tanımlayacağız.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'EĞER Fonksiyonunun Amacı ve Çalışma Prensibi',
      content:
        'EĞER fonksiyonu, bir koşulu test eden ve sonucun DOĞRU veya YANLIŞ olmasına göre iki farklı değerden birini üreten temel mantık fonksiyonudur. İnsanların "Şu olursa bunu yap, olmazsa şunu yap" şeklindeki karar alma mantığını Excel hücrelerine uyarlar.',
      keyTakeaway: 'Belirtilen koşula bağlı olarak iki farklı sonuçtan birini üretir.',
      audioNarration:
        'EĞER fonksiyonu, belirlediğimiz bir koşulu sınar. Koşul doğruysa ilk belirttiğimiz sonucu, yanlışsa ikinci sonucu hücreye yazdırır.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Fonksiyonun Sözdizimi ve Parametreleri',
      content:
        '=EĞER(mantıksal_test; [eğer_doğruysa_değer]; [eğer_yanlışsa_değer])\n\n1. Mantıksal Test: Sınanan durum (örneğin B2>=50).\n2. Eğer Doğruysa Değer: Test doğru çıktığında gösterilecek metin veya sayı ("Geçti").\n3. Eğer Yanlışsa Değer: Test sağlanmadığında gösterilecek sonuç ("Kaldı").',
      keyTakeaway: '3 parametre: Mantıksal test, doğruysa sonuç, yanlışsa sonuç.',
      audioNarration:
        'Sözdiziminde üç parametre bulunur: İlk olarak şartımızı yazarız, noktalı virgül koyup şart tutarsa ne yazacağını, tekrar noktalı virgül koyup tutmazsa ne yazacağını belirtiriz.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Mantıksal Karşılaştırma İşleçleri',
      content:
        'EĞER içinde sıkça kullanılan işleçler:\n• = (Eşittir)\n• > (Büyüktür)\n• < (Küçüktür)\n• >= (Büyük veya Eşittir)\n• <= (Küçük veya Eşittir)\n• <> (Eşit Değildir)\nMetin çıktıları formül içinde çift tırnak ("...") arasına alınmalıdır.',
      keyTakeaway: 'Sayısal karşılaştırmalarda >=, metin çıktılarında çift tırnak kullanılır.',
      audioNarration:
        'Karşılaştırma yaparken büyüktür, küçüktür veya eşit değildir gibi işleçler kullanırız. Formül içinde metin yazarken mutlaka çift tırnak işaretini unutmamalıyız.',
    },
  ],

  solvedExample: {
    title: 'Öğrenci Geçti / Kaldı Hesaplaması',
    scenario:
      'Öğretmen, B2 hücresindeki sınav notu 50 veya üzerindeyse C2 hücresine "Geçti", 50\'nin altındaysa "Kaldı" yazdırmak istemektedir.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'Sonucun yazılacağı C2 hücresini seçin.',
        tip: 'Sonuç hücresine tıklayın.',
      },
      {
        stepNumber: 2,
        instruction: '=EĞER(B2>=50;"Geçti";"Kaldı") formülünü yazıp Enter tuşuna basın.',
        tip: 'İngilizce Excel kullanıyorsanız =IF(B2>=50;"Geçti";"Kaldı") yazabilirsiniz.',
      },
    ],
    explanation:
      'B2 hücresindeki 75 değeri 50\'den büyük olduğu için koşul DOĞRU sonucunu verir ve formül hücreye "Geçti" metnini yazar.',
    audioNarration:
      'Örneğimizde B2 hücresindeki 75 notu 50 sınırını aştığı için EĞER fonksiyonu doğrudan Geçti sonucunu üretir.',
  },

  realWorldScenarios: [
    {
      industry: 'İnsan Kaynakları & Satış',
      role: 'Satış Operasyonları Müdürü',
      challenge: '50 kişilik satış ekibinde aylık 10.000 ₺ hedefini tutturanlara prim tanımlamak.',
      solution: '=EĞER(B2>=10000; "Prim Hak Etti"; "Hedef Altı") formülü ile tüm ekibin hak edişini otomatik belirlemek.',
      appliedBenefit: 'Manuel liste kontrolüyle saatler harcamak yerine saniyeler içinde hatasız prim listesi oluşturuldu.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Metin sonuçları tırnak içine almamak',
      cause: '=EĞER(B2>=50;Geçti;Kaldı) yazıldığında Excel Geçti kelimesini formül veya ad tanımlaması sanır.',
      solution: 'Tüm metin sabitlerini çift tırnak içine alın: "Geçti", "Kaldı".',
    },
    {
      mistake: 'Büyük eşit yerine yanlış sıra kullanmak (=>)',
      cause: '=> yazımı Excel tarafından sözdizimi hatası olarak kabul edilir.',
      solution: 'Her zaman önce büyüktür, sonra eşittir yazılmalıdır: >=.',
    },
  ],

  tasks: [
    {
      id: 'task_if_1',
      number: 1,
      title: '1. Adım: C2 Hücresini Seçin',
      instruction: 'İlk öğrenci olan Ahmet Yılmaz\'ın durumunu hesaplamak için C2 hücresine tıklayın.',
      narrationText: 'C2 hücresini seçerek formül girişine hazırlanalım.',
      targetCell: 'C2',
      hints: {
        level1: 'Tabloda Durum sütununun ilk boş hücresine tıklayın.',
        level2: 'C sütunu ile 2. satırın kesişimindeki C2 hücresine tıklayın.',
        level3: 'Doğrudan C2 hücresine tıklayın veya ad kutusuna C2 yazıp Enter\'a basın.',
      },
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'C', 2, 'Ahmet Yılmaz Durum'),
    },
    {
      id: 'task_if_2',
      number: 2,
      title: '2. Adım: EĞER Formülü ile Geçti/Kaldı Belirleyin',
      instruction:
        'C2 hücresine =EĞER(B2>=50;"Geçti";"Kaldı") yazıp Enter tuşuna basın.',
      narrationText:
        'B2 hücresindeki not 50 veya üzeriyse Geçti, aksi halde Kaldı sonucunu döndüreceğiz.',
      targetCell: 'C2',
      hints: {
        level1: 'Formüle eşittir ile başlayın ve EĞER(B2>=50;"Geçti";"Kaldı") yazın.',
        level2: 'Noktalı virgülleri ve çift tırnak işaretlerini doğru girdiğinizden emin olun.',
        level3: '=EĞER(B2>=50;"Geçti";"Kaldı") veya =IF(B2>=50;"Geçti";"Kaldı") yazıp Enter\'a basın.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'C2',
          [
            '=EĞER(B2>=50;"Geçti";"Kaldı")',
            '=IF(B2>=50;"Geçti";"Kaldı")',
            '=EĞER(B2>=50,"Geçti","Kaldı")',
            '=IF(B2>=50,"Geçti","Kaldı")',
          ],
          'Geçti'
        ),
    },
    {
      id: 'task_if_3',
      number: 3,
      title: '3. Adım: Zeynep Kaya İçin EĞER Formülünü Uygulayın',
      instruction:
        'C3 hücresini seçin ve Zeynep Kaya için =EĞER(B3>=50;"Geçti";"Kaldı") formülünü yazıp Enter\'a basın.',
      narrationText:
        'Zeynep\'in notu 42 olduğu için formülün Kaldı sonucunu ürettiğini gözlemleyin.',
      targetCell: 'C3',
      hints: {
        level1: 'C3 hücresine tıklayın ve aynı mantıkla B3 hücresini test eden formülü yazın.',
        level2: '=EĞER(B3>=50;"Geçti";"Kaldı") yazın.',
        level3: 'C3 hücresine =EĞER(B3>=50;"Geçti";"Kaldı") yazıp Enter\'a basın. Sonuç "Kaldı" olmalıdır.',
      },
      validate: (sheet) =>
        validateFormulaGeneral(
          sheet,
          'C3',
          [
            '=EĞER(B3>=50;"Geçti";"Kaldı")',
            '=IF(B3>=50;"Geçti";"Kaldı")',
            '=EĞER(B3>=50,"Geçti","Kaldı")',
            '=IF(B3>=50,"Geçti","Kaldı")',
          ],
          'Kaldı'
        ),
    },
  ],

  independentPractice: {
    goal: 'Satış temsilcisi Burak Can\'ın aylık satışı (B2) 10.000 ₺ ve üzeriyse C2 hücresine "Prim Var", değilse "Prim Yok" yazdırın.',
    scenario:
      'Şirketinizin prim politikasında 10.000 ₺ barajı bulunmaktadır. EĞER formülü ile Burak Can\'ın prim hak edip etmediğini hesaplayın.',
    narrationText:
      'Şimdi bağımsız uygulama adımındasınız. Satış tablosunda EĞER formülünü kendi başınıza kurun.',
    workbookFactory: createIfPracticeWorkbook,
    targetCell: 'C2',
    hints: [
      'C2 hücresini seçin.',
      'B2 hücresini 10000 ile karşılaştırın: B2>=10000',
      '=EĞER(B2>=10000;"Prim Var";"Prim Yok") yazıp Enter\'a basın.',
    ],
    validate: (sheet) =>
      validateFormulaGeneral(
        sheet,
        'C2',
        [
          '=EĞER(B2>=10000;"Prim Var";"Prim Yok")',
          '=IF(B2>=10000;"Prim Var";"Prim Yok")',
          '=EĞER(B2>=10000,"Prim Var","Prim Yok")',
          '=IF(B2>=10000,"Prim Var","Prim Yok")',
        ],
        'Prim Var'
      ),
  },

  quizQuestions: [
    {
      id: 'q_if_1',
      question: '=EĞER(A1>100;"Pahalı";"Uygun") formülünde A1 hücresi 120 ise hücrede ne görünür?',
      options: ['Pahalı', 'Uygun', 'DOĞRU', '120'],
      correctIndex: 0,
      explanation: '120 > 100 koşulu DOĞRU olduğu için birinci sonuç olan "Pahalı" döndürülür.',
    },
    {
      id: 'q_if_2',
      question: 'EĞER fonksiyonunda formül içine metin yazarken neden çift tırnak ("...") kullanılır?',
      options: [
        'Excel metinleri hücre adı veya fonksiyon zannetmesin diye',
        'Yalnızca görsel olarak daha güzel görünsün diye',
        'Metinlerin rengini mavi yapmak için',
        'Tırnak işareti zorunlu değildir',
      ],
      correctIndex: 0,
      explanation: 'Excel, tırnaksız ifadeleri tanımlı ad veya fonksiyon olarak yorumlamaya çalışır ve #AD? hatası verir.',
    },
  ],

  initialWorkbookFactory: createIfInitialWorkbook,
};
