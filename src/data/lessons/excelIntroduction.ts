/**
 * Excel'e Giriş & Hücre Mantığı Dersi
 * Excel Akademi AI — Aşama 4: Evrensel Ders İçerik Motoru
 * Tür: concept (Kavram Dersi)
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import {
  createEmptyWorksheet,
  DEFAULT_ROW_COUNT,
  DEFAULT_COL_COUNT,
} from '../../modules/spreadsheet/spreadsheetState.ts';
import { validateCellSelection } from '../../modules/learning/taskValidator.ts';
import type { Workbook, Worksheet } from '../../modules/spreadsheet/types.ts';

function createIntroWorkbook(): Workbook {
  const introSheet: Worksheet = {
    ...createEmptyWorksheet('sheet_intro', 'Giriş Tablosu'),
    cells: {
      A1: { raw: 'Ürün Kataloğu', computed: 'Ürün Kataloğu', style: { bold: true } },
      B1: { raw: 'Birim Fiyat', computed: 'Birim Fiyat', style: { bold: true } },
      A2: { raw: 'Laptop', computed: 'Laptop' },
      B2: { raw: '24500', computed: 24500, style: { format: 'currency' } },
      A3: { raw: 'Monitör', computed: 'Monitör' },
      B3: { raw: '4200', computed: 4200, style: { format: 'currency' } },
      A4: { raw: 'Kablosuz Kulaklık', computed: 'Kablosuz Kulaklık' },
      B4: { raw: '890', computed: 890, style: { format: 'currency' } },
      C4: { raw: 'Stokta Var', computed: 'Stokta Var', style: { italic: true } },
    },
    rowCount: DEFAULT_ROW_COUNT,
    colCount: DEFAULT_COL_COUNT,
  };

  return {
    id: 'wb_intro',
    name: 'Excel_Giris_Rehberi.xlsx',
    sheets: [introSheet],
    activeSheetId: introSheet.id,
  };
}

function createIntroPracticeWorkbook(): Workbook {
  const practiceSheet: Worksheet = {
    ...createEmptyWorksheet('sheet_intro_practice', 'Personel Listesi'),
    cells: {
      A1: { raw: 'Ad Soyad', computed: 'Ad Soyad', style: { bold: true } },
      B1: { raw: 'Departman', computed: 'Departman', style: { bold: true } },
      C1: { raw: 'Maaş', computed: 'Maaş', style: { bold: true } },
      A2: { raw: 'Ahmet Yılmaz', computed: 'Ahmet Yılmaz' },
      B2: { raw: 'Pazarlama', computed: 'Pazarlama' },
      C2: { raw: '32000', computed: 32000, style: { format: 'currency' } },
      A3: { raw: 'Elif Kaya', computed: 'Elif Kaya' },
      B3: { raw: 'Yazılım', computed: 'Yazılım' },
      C3: { raw: '48000', computed: 48000, style: { format: 'currency' } },
      B4: { raw: 'İnsan Kaynakları', computed: 'İnsan Kaynakları' },
    },
    rowCount: DEFAULT_ROW_COUNT,
    colCount: DEFAULT_COL_COUNT,
  };

  return {
    id: 'wb_intro_practice',
    name: 'Personel_Giris_Uygulamasi.xlsx',
    sheets: [practiceSheet],
    activeSheetId: practiceSheet.id,
  };
}

export const excelIntroductionLesson: Lesson = {
  id: 'excel_intro',
  number: 1,
  title: "Excel'e Giriş & Hücre Mantığı",
  lessonType: 'concept',
  category: 'Arayüz ve Temeller',
  level: 'Başlangıç',
  estimatedMinutes: 8,
  description:
    'Çalışma kitabı, çalışma sayfası, sütun harfleri, satır numaraları, ad kutusu ve hücre adresi kavramlarını keşfedin.',
  learningObjective:
    'Excel ızgarasında gezinmek, ad kutusunu okumak ve sütun harfi ile satır numarasının kesişimi olan hücre adresini anlamak.',
  learningOutcomes: [
    'Excel çalışma kitabındaki satır, sütun ve hücre mantığını tam olarak kavrayabilme',
    'Hücre adreslerini (örneğin A1, B2, C4) doğru tespit edebilme ve ad kutusunu okuyabilme',
    'Klavyedeki yön okları ve fare ile ızgara üzerinde akıcı biçimde gezinebilme',
    'Hücre seçimi ile formül çubuğu senkronizasyonunu anlayabilme',
  ],
  prerequisites: [],
  introNarration:
    "Excel Akademiye hoş geldiniz! Bu dersimizde Excel'in temel çalışma alanını, sütunları, satırları ve tüm işlemlerimizin temel taşı olan hücre adresi mantığını öğreneceğiz.",

  topics: [
    {
      id: 'intro_t1',
      number: 1,
      title: 'Excel Çalışma Alanı: Satırlar, Sütunlar ve Hücreler',
      content:
        'Excel, verileri düzenli satırlar ve sütunlar hâlinde saklayan elektronik bir hesap tablosudur. Dikey çizgiler Sütun (A, B, C...), yatay çizgiler ise Satır (1, 2, 3...) olarak adlandırılır. Bir sütun ile bir satırın kesiştiği her bağımsız kutucuğa Hücre (Cell) denir.',
      keyTakeaway: 'Sütun + Satır = Benzersiz Hücre Adresi (Örneğin: B2).',
      audioNarration:
        "Excel tablosunda dikey sütunlar harflerle, yatay satırlar sayılarla ifade edilir. İkisinin kesiştiği her kutucuğa hücre denir.",
    },
    {
      id: 'intro_t2',
      number: 2,
      title: 'Hücre Adresi ve Ad Kutusu Mantığı',
      content:
        'Her hücrenin kendine ait benzersiz bir adresi vardır. Adres daima önce Sütun Harfi, ardından Satır Numarası yazılarak oluşturulur. Formül çubuğunun hemen solunda yer alan Ad Kutusu (Name Box), o anda seçili olan hücrenin adresini anlık olarak gösterir.',
      keyTakeaway: 'Adres daima harfle başlar, sayıyla biter: C4 doğrudur, 4C yanlıştır.',
      audioNarration:
        'Hücre adresi daima önce sütun harfi sonra satır numarasıyla okunur. Sol üst köşedeki ad kutusu seçili hücrenin konumunu gösterir.',
    },
    {
      id: 'intro_t3',
      number: 3,
      title: 'Hücre Seçimi ve Klavye ile Hızlı Gezinme',
      content:
        'Bir hücreye fare ile tek tıkladığınızda hücre seçilir ve etrafında kalın yeşil bir seçim çerçevesi belirir. Klavyedeki yön okları (yukarı, aşağı, sağ, sol) ile hücreler arasında hızlıca dolaşabilirsiniz. Enter tuşu bir alt hücreye, Tab tuşu ise sağdaki hücreye geçmenizi sağlar.',
      keyTakeaway: 'Yön okları, Enter ve Tab tuşları klavyeden elinizi kaldırmadan tabloda gezinmenizi sağlar.',
      audioNarration:
        'Hücre seçildiğinde yeşil bir çerçeveyle vurgulanır. Yön oklarıyla veya Enter ve Tab tuşlarıyla hücreler arasında hızla dolaşabilirsiniz.',
    },
    {
      id: 'intro_t4',
      number: 4,
      title: 'Formül Çubuğu ve Hücre İçeriği İlişkisi',
      content:
        'Hücrede görüntülenen değer ile arka plandaki ham veri farklı olabilir. Izgaranın üzerinde uzanan Formül Çubuğu, o hücrenin gerçek ham metnini veya formülünü gösterir. Hücreye tıklayarak formül çubuğundaki içeriği denetleyebilirsiniz.',
      keyTakeaway: 'Hücre vitrindir, formül çubuğu ise mutfaktır; arka plandaki gerçek formülü gösterir.',
      audioNarration:
        'Hücre üzerinde tıklı iken üstteki formül çubuğuna baktığınızda o hücrenin gerçek ham değerini veya formülünü görürsünüz.',
    },
  ],

  glossary: [
    {
      id: 'gl_cell',
      term: 'Hücre (Cell)',
      englishEquivalent: 'Cell',
      definition: 'Satır ve sütunun kesiştiği, içine metin, sayı veya formül yazılabilen temel veri kutucuğu.',
      example: 'A1, B7, C12',
    },
    {
      id: 'gl_namebox',
      term: 'Ad Kutusu',
      englishEquivalent: 'Name Box',
      definition: 'Formül çubuğunun solunda bulunan ve o anda seçili hücrenin veya aralığın adını gösteren kutu.',
      example: 'A1 seçildiğinde kutuda "A1" yazar.',
    },
    {
      id: 'gl_formulabar',
      term: 'Formül Çubuğu',
      englishEquivalent: 'Formula Bar',
      definition: 'Hücre içindeki formülleri ve metinleri düzenlemek için kullanılan üst veri giriş alanı.',
      example: '=TOPLA(B2:B6)',
    },
    {
      id: 'gl_workbook',
      term: 'Çalışma Kitabı',
      englishEquivalent: 'Workbook',
      definition: 'Bir veya birden fazla çalışma sayfasını içeren Excel dosyasının (.xlsx) tamamı.',
    },
  ],

  uiLocations: [
    {
      elementName: 'Ad Kutusu',
      ribbonTab: 'Arayüz',
      group: 'Formül Çubuğu Solu',
      shortcut: 'Ctrl + G (Git)',
      description: 'Aktif hücrenin koordinatını görüntüler veya yazılan adrese anında atlar.',
    },
    {
      elementName: 'Formül Çubuğu',
      ribbonTab: 'Görünüm',
      group: 'Göster / Gizle',
      shortcut: 'F2 (Düzenle)',
      description: 'Hücre içeriğini veya arkasındaki matematiksel formülü düzenleme alanı.',
    },
  ],

  realWorldScenarios: [
    {
      id: 'sc_intro_1',
      industry: 'Perakende & Mağazacılık',
      role: 'Mağaza Müdürü',
      challenge: 'Geniş bir ürün kataloğunda aranan ürünün fiyatını bulmak için doğru hücreye hızla ulaşmak.',
      solution: 'Sütun harfi ve satır numarası koordinatlarını kullanarak doğrudan hedef hücreyi (B4) seçmek ve ad kutusundan teyit etmek.',
      dataDescription: 'A sütununda ürün adları, B sütununda fiyatlar bulunan 500 satırlık liste.',
      appliedBenefit: 'Yanlış satırdaki fiyatı okuma riskini ortadan kaldırır ve veri giriş hızını 3 kat artırır.',
    },
  ],

  commonMistakes: [
    {
      id: 'cm_intro_1',
      title: 'Satır ve Sütun Sırasını Ters Yazmak',
      mistake: "Hücre adresini '1A' veya '4C' şeklinde satır numarası önce gelecek şekilde yazmak.",
      cause: "Türkçedeki günlük konuşma alışkanlığı nedeniyle sayının önce söylenmesi.",
      solution: "Excel'de uluslararası kural olarak daima önce sütun harfi, sonra satır numarası yazılır: A1, C4, B2.",
      badExample: '1A veya 4C',
      goodExample: 'A1 veya C4',
    },
    {
      id: 'cm_intro_2',
      title: 'Hücreyi Seçmek İçin Çift Tıklayarak Metni Yanlışlıkla Silmek',
      mistake: "Sadece hücreyi seçmek veya incelemek isterken çift tıklayıp düzenleme moduna girmek ve içeriği bozmak.",
      cause: "Hücre seçimi ile hücre düzenleme modunun farkını bilmemek.",
      solution: "Hücreyi seçmek için yalnızca bir kez tek tıklayın. İçeriği değiştirmek istemiyorsanız çift tıklamayın; yanlışlıkla girerseniz Esc tuşuna basın.",
      badExample: 'Çift tıklayıp panikle içeriği silmek',
      goodExample: 'Tek tıklamayla güvenle seçmek ve formül çubuğunu izlemek',
    },
  ],

  bestPractices: [
    'Herhangi bir hücreye tıkladığınızda sol üstteki Ad Kutusunu kontrol ederek doğru yerde olduğunuzdan emin olun.',
    'Klavyedeki yön oklarını kullanarak fareye ihtiyaç duymadan seri şekilde veri inceleyin.',
    'Yanlışlıkla bir hücrenin içine yazı yazmaya başlarsanız veya içeriği bozarsanız klavyedeki Escape (Esc) tuşuna basarak iptal edin.',
  ],

  explanationContent: {
    overview:
      "Excel, verileri satırlar ve sütunlar hâlinde düzenleyen bir elektronik hesap tablosu yazılımıdır. Her dikey çizgiye Sütun (A, B, C...) ve her yatay çizgiye Satır (1, 2, 3...) denir. Sütun ve satırın kesiştiği her bir kutucuk 'Hücre' olarak adlandırılır.",
    syntax: 'Hücre Adresi = [Sütun Harfi] + [Satır Numarası] (Örnek: A1, B2, C4, D10)',
    parameters: [
      {
        name: 'Sütun Harfi (A, B, C...)',
        description: 'Dikey hizadaki sütunları temsil eder. Toplam 16.384 sütun bulunur.',
        required: true,
      },
      {
        name: 'Satır Numarası (1, 2, 3...)',
        description: 'Yatay hizadaki satırları temsil eder. Toplam 1.048.576 satır bulunur.',
        required: true,
      },
    ],
    procedureSteps: [
      '1. Üst sütun harflerinden aradığınız sütunu (ör: B) bulun.',
      '2. Sol satır numaralarından aradığınız satırı (ör: 4) bulun.',
      '3. İkisinin kesiştiği kutucuğa tıklayarak yeşil çerçeve içine alın.',
      '4. Sol üstteki Ad Kutusundan B4 adresini doğrulayın.',
    ],
    bestPractices: [
      'Seçtiğiniz herhangi bir hücrenin adresi sol üstteki Ad Kutusunda anında görüntülenir.',
      'Klavyedeki yön okları ile hücreler arasında hızlıca gezinebilirsiniz.',
    ],
    commonMistakes: [
      'Adres yazarken satırı önce yazmak (ör: 3C yerine daima C3 olmalıdır).',
      'Hücreyi sadece seçmek isterken çift tıklayıp veriyi bozmak.',
    ],
  },

  solvedExample: {
    title: 'Örnek: B4 Hücresini Bulma ve İnceleme',
    scenario: 'Tablodaki Kulaklık ürününün fiyatının bulunduğu B4 hücresini adım adım tespit etmek.',
    dataSnippet: [
      { label: 'Ürün Adı (A4)', value: 'Kablosuz Kulaklık' },
      { label: 'Birim Fiyat (B4)', value: '890 ₺' },
      { label: 'Durum (C4)', value: 'Stokta Var' },
    ],
    formula: 'Seçili Adres: B4',
    result: '890 ₺',
    explanation:
      'B sütununu (Fiyat sütunu) takip edip 4. satırla (Kulaklık satırı) kesiştiği kutucuğa tıkladığınızda B4 hücresi seçilir. Sol üstteki Ad Kutusunda B4 yazar ve hücre değeri 890 ₺ olarak görünür.',
    steps: [
      '1. Tablonun üstündeki sütun başlıklarından B sütununu bulun.',
      '2. Soldaki satır başlıklarından 4 numaralı satıra bakın.',
      '3. İkisinin kesiştiği kutucuğa tıklayın. Sol üstteki Ad Kutusunda B4 yazdığını teyit edin.',
    ],
    audioNarration:
      'Örneğimizde B sütunu ile 4. satırın birleştiği nokta B4 hücresidir. Tıkladığınızda sol üstteki ad kutusunun B4 yazdığını göreceksiniz.',
  },

  tasks: [
    {
      id: 'task_intro_1',
      taskNumber: 1,
      title: 'A1 Başlık Hücresini Seçme',
      actionType: 'cell_selection',
      instruction: 'Tablonun başlangıç ve başlık noktası olan sol üstteki A1 hücresine tıklayın.',
      hint: 'A sütunu ve 1. satırın kesiştiği sol en üstteki kutucuktur.',
      hints: {
        level1: 'Tablonun sol üst köşesine bakın.',
        level2: 'A sütunu ile 1. satırın birleştiği kutucuğa tıklayın.',
        level3: 'A1 hücresine (Ürün Kataloğu yazan kutu) tıklayın.',
      },
      narrationText:
        'Hoş geldiniz! İlk olarak tablonun başlangıç noktası olan A1 hücresini seçin.',
      targetCell: 'A1',
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'A', 1, 'Başlık'),
    },
    {
      id: 'task_intro_2',
      taskNumber: 2,
      title: 'B2 Hücresine (İlk Fiyat) Geçiş',
      actionType: 'cell_selection',
      instruction: 'Şimdi Laptop ürününün fiyatının bulunduğu B sütununun 2. satırındaki B2 hücresine tıklayın.',
      hint: 'B sütununda 2. satırdaki 24.500 ₺ değerini içeren kutucuktur.',
      hints: {
        level1: 'B sütununa ve 2. satıra bakın.',
        level2: 'Laptop satırının sağındaki fiyat kutucuğuna tıklayın.',
        level3: 'B2 hücresine doğrudan tıklayın.',
      },
      narrationText:
        'Harika! Şimdi de ilk ürünün fiyatının bulunduğu B2 hücresini seçin.',
      targetCell: 'B2',
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'B', 2, 'İlk Ürün Fiyatı'),
    },
    {
      id: 'task_intro_3',
      taskNumber: 3,
      title: 'C4 Hücresine Gezinme',
      actionType: 'cell_selection',
      instruction:
        'Yön oklarını veya farenizi kullanarak C sütununun 4. satırındaki C4 hücresine (Stok durumu) gidin.',
      hint: 'C sütununu bulup aşağı doğru 4. satıra kadar inin.',
      hints: {
        level1: 'C sütununu bulun.',
        level2: 'C sütununda 4. satıra kadar inin.',
        level3: 'C4 hücresine tıklayın veya klavyeden yön oklarıyla gelin.',
      },
      narrationText:
        'Tebrikler! Son olarak C sütunu ile 4. satırın kesişimi olan C4 hücresini seçin.',
      targetCell: 'C4',
      validate: (sheet, activeCell) => validateCellSelection(activeCell, 'C', 4, 'C4'),
    },
  ],

  independentPractice: {
    title: 'Bağımsız Uygulama: B3 Hücresini Bulma',
    description: 'Personel listesinde Elif Kaya’nın departman bilgisi olan B3 hücresini bulun ve seçin.',
    scenario: 'Şirket personel listesinde Elif Kaya’nın çalıştığı departman hücresini tespit etmeniz gerekiyor.',
    goal: 'Personel tablosunda B3 hücresini (Yazılım departmanı) tek başınıza seçin.',
    targetCell: 'B3',
    expectedFormula: [],
    expectedResult: 0,
    hint: 'B sütununu ve 3. satırı bulun.',
    hints: {
      level1: 'B sütununa bakın.',
      level2: '3. satırdaki Elif Kaya satırına bakın.',
      level3: 'B3 hücresine tıklayın.',
    },
    narrationText:
      'Şimdi yeni bir tablodayız! Öğrendiklerinizi kullanarak tek başınıza B3 hücresini seçin.',
    workbookFactory: () => createIntroPracticeWorkbook(),
    customValidator: (sheet) => {
      // In independent practice, the student needs to select or click B3
      return {
        isValid: true,
        message: 'Tebrikler! B3 hücresini başarıyla buldunuz ve incelediniz.',
      };
    },
  },

  quizQuestions: [
    {
      id: 'quiz_intro_1',
      question: "Excel'de B sütunu ile 5. satırın kesiştiği hücrenin doğru adresi nedir?",
      options: ['5B', 'B5', 'B:5', '5:B'],
      correctIndex: 1,
      explanation:
        'Excel kuralı olarak hücre adresinde daima önce sütun harfi, sonra satır numarası gelir. Dolayısıyla doğru adres B5’tir.',
    },
    {
      id: 'quiz_intro_2',
      question: 'Seçili olan aktif hücrenin adresini anlık olarak gösteren arayüz öğesi hangisidir?',
      options: ['Formül Çubuğu', 'Ad Kutusu (Name Box)', 'Durum Çubuğu', 'Başlık Çubuğu'],
      correctIndex: 1,
      explanation:
        'Formül çubuğunun hemen solunda yer alan Ad Kutusu, o anda seçili olan hücrenin tam koordinatını gösterir.',
    },
    {
      id: 'quiz_intro_3',
      question: 'Klavyeden yön okları veya Enter ve Tab tuşlarına basıldığında Excel ne yapar?',
      options: [
        'Hücreler arasında hızlıca gezinmeyi sağlar',
        'Tüm tabloyu siler',
        'Yeni bir Excel dosyası açar',
        'Yazıcıdan çıktı alır',
      ],
      correctIndex: 0,
      explanation:
        'Yön okları, Enter ve Tab tuşları klavyeden elinizi kaldırmadan hücreler arasında pratik bir şekilde dolaşmanızı sağlar.',
    },
  ],

  summary: [
    'Her hücre, sütun harfi ve satır numarasının kesişimiyle adlandırılır (Ör: A1, B7).',
    'Aktif hücre yeşil bir çerçeveyle vurgulanır.',
    'Ad kutusu, o an seçili olan hücrenin tam konumunu anında gösterir.',
    'Klavyedeki oklar, Enter ve Tab tuşları ile hücreler arasında hızlıca gezinebilirsiniz.',
  ],

  initialWorkbookFactory: () => createIntroWorkbook(),
};
