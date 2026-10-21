/**
 * Grafik Oluşturma Dersi - Excel Akademi AI
 * Aşama 5: Temel Grafikler (Sütun, Çizgi, Pasta) ve Görsel Veri Analizi
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateChartExists,
  validateRangeSelection,
} from '../../modules/learning/taskValidator.ts';

export function createChartInitialWorkbook(): Workbook {
  return {
    id: 'wb_chart_lesson',
    name: 'Aylik_Satis_Grafigi.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Aylık Satışlar',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Ay', computed: 'Ay', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Ciro (₺)', computed: 'Ciro (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Ocak', computed: 'Ocak' },
          B2: { raw: '12000', computed: 12000, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Şubat', computed: 'Şubat' },
          B3: { raw: '18000', computed: 18000, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Mart', computed: 'Mart' },
          B4: { raw: '15000', computed: 15000, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'Nisan', computed: 'Nisan' },
          B5: { raw: '24000', computed: 24000, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export function createChartPracticeWorkbook(): Workbook {
  return {
    id: 'wb_chart_practice',
    name: 'Gider_Dagilimi_Pasta.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Giderler',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'Kategori', computed: 'Kategori', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Harcama (₺)', computed: 'Harcama (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'Kira', computed: 'Kira' },
          B2: { raw: '25000', computed: 25000, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'Maaşlar', computed: 'Maaşlar' },
          B3: { raw: '60000', computed: 60000, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'Pazarlama', computed: 'Pazarlama' },
          B4: { raw: '15000', computed: 15000, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export const chartCreationLesson: Lesson = {
  id: 'chart_creation_lesson',
  number: 10,
  title: 'Temel Grafikler: Sütun, Çizgi ve Pasta',
  lessonType: 'visualization',
  category: 'Veri Görselleştirme',
  level: 'Başlangıç-Orta',
  estimatedMinutes: 16,
  description:
    'Rakamlarla dolu tablolarınızı etkileyici sütun, çizgi ve pasta grafiklerine dönüştürerek canlı görsel raporlar oluşturmayı öğrenin.',
  learningObjective:
    'Grafik türlerinin amacını (karşılaştırma, eğilim, oran), veri ve kategori aralıklarının seçilmesini ve hücre değiştikçe grafiğin otomatik güncellenmesini kavramak.',
  learningOutcomes: [
    'Veriye uygun grafik türünü (Sütun, Çizgi veya Pasta) doğru seçebilme',
    'Kategori ve değer aralıklarını tanımlayarak canlı SVG grafiği üretebilme',
    'Tablodaki hücre değeri değiştiğinde grafiğin otomatik tepki vermesini gözlemleme',
  ],
  prerequisites: [
    {
      id: 'excel_intro',
      title: "Excel'e Giriş & Hücre Mantığı",
    },
    {
      id: 'sum_function',
      title: 'TOPLA (SUM) Fonksiyonu',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde "bir görsel bin söze bedeldir" ilkesiyle hareket edeceğiz. Tablomuzdaki aylık ciro rakamlarını renkli ve dinamik bir Sütun Grafiğine dönüştürerek yöneticilerin ve müşterilerin ilk bakışta anlayabileceği profesyonel bir sunum hazırlayacağız.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'Grafiklerin Gücü ve Grafik Türleri',
      content:
        'Grafikler, sayı yığınlarını tek bakışta anlaşılan hikayelere dönüştürür.\n• Sütun Grafiği: Farklı kategorileri veya aylar arası ciroları karşılaştırmak için idealdir.\n• Çizgi Grafiği: Zaman içindeki yükseliş ve düşüş eğilimlerini (trend) gösterir.\n• Pasta Grafiği: Bir bütünün parçalara dağılımını (örneğin bütçe payları) sergiler.',
      keyTakeaway: 'Sütun karşılaştırır, çizgi eğilimi izler, pasta dağılımı gösterir.',
      audioNarration:
        'Doğru grafik türü seçmek iletişimin anahtarıdır. Farklı ayları kıyaslamak için sütun, zaman trendi için çizgi, pasta grafiğini ise payları göstermek için kullanırız.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Kategori ve Değer Aralığı Mantığı',
      content:
        'Her grafiğin iki ana bileşeni vardır:\n1. Değer Aralığı (Y Ekseni): Grafik barlarının yüksekliğini belirleyen sayısal hücreler (örneğin B2:B5).\n2. Kategori Aralığı (X Ekseni): Barların altına yazılan isimler veya aylar (örneğin A2:A5).',
      keyTakeaway: 'Sayılar değer aralığını, metin etiketleri kategori aralığını oluşturur.',
      audioNarration:
        'Grafik oluştururken sayısal sütunumuzu değer aralığı, ay isimlerini içeren sütunumuzu ise kategori aralığı olarak belirleriz.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Dinamik Canlı Güncellenme',
      content:
        'Excel grafikleri statik bir resim değildir; doğrudan tablodaki hücrelere bağlıdır. Tablodaki bir satışı 12.000\'den 30.000\'e çıkardığınızda, altındaki grafik sütunu anında uzar ve rengini korur.',
      keyTakeaway: 'Hücre verisi değiştiğinde grafik eşzamanlı olarak yeniden çizilir.',
      audioNarration:
        'Excel grafiklerinin en güzel tarafı canlı olmasıdır. Tablonuzdaki bir sayıyı değiştirdiğiniz anda grafik kendini anında yeniler.',
    },
  ],

  solvedExample: {
    title: 'Aylık Satış Sütun Grafiği Oluşturma',
    scenario:
      'A2:B5 aralığındaki 4 aylık satış verilerini gösteren bir Sütun Grafiği eklemek istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'Üst araç çubuğundaki "Grafik Ekle" butonuna tıklayın.',
        tip: 'Grafik oluşturma penceresi açılacaktır.',
      },
      {
        stepNumber: 2,
        instruction: 'Grafik Türünü "Sütun Grafiği", Kategori Aralığını A2:A5 ve Değer Aralığını B2:B5 seçip "Grafik Oluştur" butonuna basın.',
        tip: 'Çalışma sayfasının altında interaktif SVG grafiğiniz belirecektir.',
      },
    ],
    explanation:
      'Grafik Ocak (12.000), Şubat (18.000), Mart (15.000) ve Nisan (24.000) değerlerini mavi sütunlarla ölçeklendirerek görselleştirir.',
    audioNarration:
      'Grafik Ekle penceresini açıp sütun türünü seçtiğimizde tüm aylar için dinamik bir görsel şölen hazırlanır.',
  },

  realWorldScenarios: [
    {
      industry: 'Yönetim Danışmanlığı & C-Level Raporlama',
      role: 'İş Geliştirme Direktörü',
      challenge: 'Yönetim kurulu toplantısında 15 sayfalık finansal tabloyu 10 dakikada karar vericilere aktarmak.',
      solution: 'Özet çeyrek rakamlarını Sütun ve Pasta grafikleriyle donatılmış interaktif bir dashboard\'a çevirmek.',
      appliedBenefit: 'Yönetim kurulu rakamları anında kavradı ve bütçe onayı tek toplantıda alındı.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Değer aralığına metin sütununu seçmek',
      cause: 'Sayı yerine isim sütunu seçildiğinde barlar çizilemez ve yükseklik 0 kalır.',
      solution: 'Değer aralığına daima sayı içeren sütunları (B2:B5) bağlayın.',
    },
    {
      mistake: 'Pasta grafiğinde 15\'ten fazla dilim kullanmak',
      cause: 'Çok fazla dilim pasta grafiğini okunmaz ve karmaşık hale getirir.',
      solution: 'Pasta grafiğini yalnızca 3 ila 7 temel kategori arasındaki dağılımlar için tercih edin.',
    },
  ],

  tasks: [
    {
      id: 'task_chart_1',
      number: 1,
      title: '1. Adım: Veri Aralığını Seçin',
      instruction: 'Grafiğe aktarmak istediğimiz verileri seçmek için A2:B5 aralığını işaretleyin.',
      narrationText: 'A2 ile B5 arasındaki hücre alanını seçin.',
      targetCell: 'A2',
      hints: {
        level1: 'A2 hücresinden B5 hücresine kadar farenizle sürükleyin.',
        level2: 'A2:B5 aralığını seçin.',
        level3: 'A2\'ye tıklayın, Shift tuşuna basılı tutarak B5\'e tıklayın.',
      },
      validate: (sheet, activeCell, isEditing, selectionRange) =>
        validateRangeSelection(selectionRange || null, activeCell, 'A2:B5'),
    },
    {
      id: 'task_chart_2',
      number: 2,
      title: '2. Adım: Canlı Sütun Grafiği Ekleyin',
      instruction:
        'Araç çubuğundaki "Grafik Ekle" butonuna tıklayın, tür olarak "Sütun Grafiği" (Column), Değer Aralığı olarak B2:B5 belirleyip "Grafik Oluştur" butonuna basın.',
      narrationText:
        'Grafik Ekle butonuna basarak sayfamıza dinamik sütun grafiğini ekleyelim.',
      targetCell: 'B2',
      hints: {
        level1: 'Araç çubuğunda "Grafik Ekle" butonuna tıklayın.',
        level2: 'Açılan pencerede Grafik Türü: Sütun Grafiği, Değer Aralığı: B2:B5 olsun.',
        level3: 'Grafik Oluştur butonuna tıklayın. Sayfanın altında grafik görünecektir.',
      },
      validate: (sheet) => validateChartExists(sheet, 'column', 'B2:B5'),
    },
  ],

  independentPractice: {
    goal: 'Giderler tablosunda harcama paylarını göstermek için Kategori (A2:A4) ve Harcama (B2:B4) verilerini içeren bir "Pasta Grafiği" (Pie) oluşturun.',
    scenario:
      'Şirketin bütçe toplantısı için Kira, Maaş ve Pazarlama giderlerinin orantısal payını pasta grafiği ile sergileyin.',
    narrationText:
      'Bağımsız alıştırmada Pasta Grafiği oluşturarak harcama paylarını görselleştirin.',
    workbookFactory: createChartPracticeWorkbook,
    targetCell: 'B2',
    hints: [
      'Araç çubuğundaki Grafik Ekle butonuna tıklayın.',
      'Grafik Türü alanından "Pasta Grafiği (Pie)" seçeneğini seçin.',
      'Kategori Aralığı: A2:A4, Değer Aralığı: B2:B4 belirleyip Grafik Oluştur\'a basın.',
    ],
    validate: (sheet) => validateChartExists(sheet, 'pie', 'B2:B4'),
  },

  quizQuestions: [
    {
      id: 'q_chart_1',
      question: 'Bir bütçenin veya harcamanın oransal yüzdelerini (paylarını) göstermek için en uygun grafik türü hangisidir?',
      options: ['Pasta Grafiği (Pie Chart)', 'Çizgi Grafiği', 'Radar Grafiği', 'Dağılım Grafiği'],
      correctIndex: 0,
      explanation: 'Pasta grafikleri, bir bütünün parçalara yüzdesel dağılımını sergilemek için standarttır.',
    },
    {
      id: 'q_chart_2',
      question: 'Tablodaki bir sayı değiştiğinde Excel grafiği ne yapar?',
      options: [
        'Grafik kendini anında otomatik olarak yeni sayıya göre günceller',
        'Grafik silinir',
        'Grafiği silip baştan çizmek gerekir',
        'Hata uyarısı verir',
      ],
      correctIndex: 0,
      explanation: 'Excel grafikleri tablodaki hücrelerle dinamik bağlıdır ve hücre değiştikçe anında güncellenir.',
    },
  ],

  initialWorkbookFactory: createChartInitialWorkbook,
};
