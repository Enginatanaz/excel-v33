/**
 * Veri Filtreleme Dersi - Excel Akademi AI
 * Aşama 5: Filtreleme, Gizlenen Satırlar ve Tablo Yönetimi
 */

import type { Lesson } from '../../modules/learning/lessonTypes.ts';
import type { Workbook } from '../../modules/spreadsheet/types.ts';
import {
  validateTableFiltered,
} from '../../modules/learning/taskValidator.ts';

export function createFilterInitialWorkbook(): Workbook {
  return {
    id: 'wb_filter_lesson',
    name: 'Siparis_Filtreleme_Tablosu.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Siparişler',
        rowCount: 15,
        colCount: 6,
        cells: {
          A1: { raw: 'Sipariş No', computed: 'Sipariş No', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Müşteri', computed: 'Müşteri', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Şehir', computed: 'Şehir', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          D1: { raw: 'Tutar (₺)', computed: 'Tutar (₺)', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'SP-101', computed: 'SP-101' },
          B2: { raw: 'Atlas Ltd.', computed: 'Atlas Ltd.' },
          C2: { raw: 'İstanbul', computed: 'İstanbul' },
          D2: { raw: '2500', computed: 2500, style: { align: 'right', format: 'currency' } },
          A3: { raw: 'SP-102', computed: 'SP-102' },
          B3: { raw: 'Birlik A.Ş.', computed: 'Birlik A.Ş.' },
          C3: { raw: 'Ankara', computed: 'Ankara' },
          D3: { raw: '1400', computed: 1400, style: { align: 'right', format: 'currency' } },
          A4: { raw: 'SP-103', computed: 'SP-103' },
          B4: { raw: 'Cihan Dış Tic.', computed: 'Cihan Dış Tic.' },
          C4: { raw: 'İstanbul', computed: 'İstanbul' },
          D4: { raw: '3800', computed: 3800, style: { align: 'right', format: 'currency' } },
          A5: { raw: 'SP-104', computed: 'SP-104' },
          B5: { raw: 'Doğan Lojistik', computed: 'Doğan Lojistik' },
          C5: { raw: 'İzmir', computed: 'İzmir' },
          D5: { raw: '900', computed: 900, style: { align: 'right', format: 'currency' } },
        },
      },
    ],
  };
}

export function createFilterPracticeWorkbook(): Workbook {
  return {
    id: 'wb_filter_practice',
    name: 'Depo_Sevkiyat_Filtre.xlsx',
    activeSheetId: 'sheet1',
    sheets: [
      {
        id: 'sheet1',
        name: 'Sevkiyatlar',
        rowCount: 12,
        colCount: 6,
        cells: {
          A1: { raw: 'İrsaliye', computed: 'İrsaliye', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          B1: { raw: 'Bölge', computed: 'Bölge', style: { bold: true, align: 'left', bgColor: '#f1f5f9' } },
          C1: { raw: 'Koli Adedi', computed: 'Koli Adedi', style: { bold: true, align: 'right', bgColor: '#f1f5f9' } },
          A2: { raw: 'IRS-1', computed: 'IRS-1' },
          B2: { raw: 'Marmara', computed: 'Marmara' },
          C2: { raw: '45', computed: 45, style: { align: 'right' } },
          A3: { raw: 'IRS-2', computed: 'IRS-2' },
          B3: { raw: 'Ege', computed: 'Ege' },
          C3: { raw: '20', computed: 20, style: { align: 'right' } },
          A4: { raw: 'IRS-3', computed: 'IRS-3' },
          B4: { raw: 'Marmara', computed: 'Marmara' },
          C4: { raw: '80', computed: 80, style: { align: 'right' } },
        },
      },
    ],
  };
}

export const dataFilterLesson: Lesson = {
  id: 'data_filter_lesson',
  number: 9,
  title: 'Veri Filtreleme ve Tablo Yönetimi',
  lessonType: 'data_management',
  category: 'Veri Yönetimi & Düzenleme',
  level: 'Başlangıç-Orta',
  estimatedMinutes: 14,
  description:
    'Binlerce satırlık listelerde yalnızca ilgilendiğiniz verileri ekranda tutup diğer satırları geçici olarak gizleyen Otomatik Filtreleme özelliğini öğrenin.',
  learningObjective:
    'Filtreleme ile satır silme arasındaki farkı kavramak, sütun bazlı koşullu filtreler uygulamak ve filtreyi temizleme adımlarını öğrenmek.',
  learningOutcomes: [
    'Çalışma tablosunda Filtre düğmelerini aktif ve pasif hale getirebilme',
    'Metin ve sayısal filtreleme koşulları belirleyerek ilgisiz satırları gizleyebilme',
    'Filtrelenen satırların silinmediğini, yalnızca geçici olarak gizlendiğini anlama',
  ],
  prerequisites: [
    {
      id: 'data_sort_lesson',
      title: 'Veri Sıralama (A-Z ve Z-A)',
    },
  ],

  introNarration:
    'Excel Akademi AI’ya hoş geldiniz! Bu dersimizde büyük tablolarda aradığımız spesifik bilgilere odaklanmamızı sağlayan Filtreleme özelliğini inceleyeceğiz. Şehir sütunundan sadece İstanbul siparişlerini seçerek diğer satırları nasıl güvenle gizleyeceğimizi göreceğiz.',

  topics: [
    {
      id: 'topic_1',
      number: 1,
      title: 'Filtreleme Nedir ve Neden Kullanılır?',
      content:
        'Filtreleme, tablonuzdaki satırları belirlediğiniz ölçütlere göre süzer. Filtre dışı kalan satırlar asla SİLİNMEZ; sadece geçici olarak ekrandan gizlenir. Filtreyi temizlediğinizde veya kaldırdığınızda tüm satırlar eksiksiz olarak geri gelir.',
      keyTakeaway: 'Filtreleme satırları silmez, yalnızca geçici olarak gizler.',
      audioNarration:
        'Filtreleme özelliğinin en büyük gücü verilerinizi korumasıdır. Filtre uygulandığında ilgisiz satırlar silinmez, sadece göz önünden geçici olarak kaldırılır.',
    },
    {
      id: 'topic_2',
      number: 2,
      title: 'Filtre Modunu Etkinleştirme',
      content:
        'Araç çubuğundaki "Filtrele" butonuna tıklandığında, tablonun en üst başlık satırındaki her bir sütun hücresinin yanına küçük bir aşağı ok simgesi eklenir. Bu simgeye tıklanarak o sütun için özel filtreleme kuralları belirlenir.',
      keyTakeaway: 'Başlıklardaki aşağı ok simgeleri filtrenin aktif olduğunu gösterir.',
      audioNarration:
        'Filtrele butonuna bastığınızda başlıkların yanında küçük açılır menü okları belirir. Buradan istediğiniz kriteri seçebilirsiniz.',
    },
    {
      id: 'topic_3',
      number: 3,
      title: 'Filtre Temizleme ve Tüm Verileri Gösterme',
      content:
        'Filtrelenmiş bir tabloda işlem bittikten sonra filtre simgesine tıklanıp "Filtreyi Temizle" seçildiğinde tüm satırlar yeniden görünür hale gelir. Böylece tablonun bütünlüğü her zaman korunmuş olur.',
      keyTakeaway: 'Filtre temizlendiğinde gizli satırlar anında geri gelir.',
      audioNarration:
        'İşiniz bittiğinde filtreyi temizleyerek tablonun tüm satırlarını saniyeler içinde eski haline getirebilirsiniz.',
    },
  ],

  solvedExample: {
    title: 'İstanbul Şehrindeki Siparişleri Filtreleme',
    scenario:
      'Sipariş tablosunda yalnızca Şehir sütunu "İstanbul" olan siparişleri ekranda bırakmak istiyoruz.',
    steps: [
      {
        stepNumber: 1,
        instruction: 'Üst araç çubuğundaki "Filtrele" butonuna tıklayarak filtreleri açın.',
        tip: 'Başlıklarda aşağı ok simgeleri belirecektir.',
      },
      {
        stepNumber: 2,
        instruction: 'C sütunu başlığındaki filtre simgesine tıklayıp "İstanbul" değerini seçin.',
        tip: 'Sadece İstanbul satırları kalacak, Ankara ve İzmir satırları gizlenecektir.',
      },
    ],
    explanation:
      'SP-101 ve SP-103 siparişleri ekranda kalırken, SP-102 ve SP-104 satırları gizlenir.',
    audioNarration:
      'Filtremizi uyguladığımızda Ankara ve İzmir satırları gizlenir, sadece İstanbul siparişleri görüntülenir.',
  },

  realWorldScenarios: [
    {
      industry: 'Tedarik Zinciri & Satınalma',
      role: 'Kategori Yöneticisi',
      challenge: '15.000 satırlık hammadde listesinde teslimat tarihi gecikmiş olan siparişleri ayıklamak.',
      solution: 'Durum sütununda "Gecikmiş" filtresi uygulayarak anında tedarikçi aksiyon listesi oluşturmak.',
      appliedBenefit: 'Saatlerce sürecek manuel liste taraması 3 saniyelik bir filtreleme hamlesine dönüştü.',
    },
  ],

  commonMistakes: [
    {
      mistake: 'Filtrelenmiş satırları silinmiş zannedip yeniden eklemek',
      cause: 'Filtre uygulandığında satır numaralarının atladığını (1, 2, 4) gören kullanıcılar satırların kaybolduğunu sanabilir.',
      solution: 'Satır numaralarındaki atlama gizlenen satırları gösterir; filtre temizlenince tüm veriler döner.',
    },
    {
      mistake: 'Filtre açıkken kopyalama yaparken tüm gizli satırları da almak',
      cause: 'Eski Excel sürümlerinde gizli hücreler de kopyalanabiliyordu.',
      solution: 'Modern Excel yalnızca görünür hücreleri kopyalar.',
    },
  ],

  tasks: [
    {
      id: 'task_filter_1',
      number: 1,
      title: '1. Adım: Filtreleme Modunu Etkinleştirin',
      instruction: 'Üst araç çubuğundaki "Filtrele" butonuna tıklayarak tablo başlıklarına filtre butonlarını ekleyin.',
      narrationText: 'Filtrele butonuna basarak başlıklarımıza filtre simgelerini yerleştirelim.',
      targetCell: 'A1',
      hints: {
        level1: 'Araç çubuğunda huni simgeli "Filtrele" butonunu bulun.',
        level2: 'Filtrele butonuna bir kez tıklayın.',
        level3: 'Üst şeritteki Filtrele butonuna tıklayın.',
      },
      validate: (sheet) => validateTableFiltered(sheet),
    },
    {
      id: 'task_filter_2',
      number: 2,
      title: '2. Adım: Şehir Sütununu "İstanbul" Olarak Filtreleyin',
      instruction:
        'C sütun başlığındaki açılır menü simgesine tıklayın, ölçüt olarak "İstanbul" girip filtreyi uygulayın.',
      narrationText:
        'C sütunundaki filtre menüsünü açın ve sadece İstanbul satırlarını görüntüleyecek şekilde filtreleyin.',
      targetCell: 'C1',
      hints: {
        level1: 'C sütunu başlığındaki filtre okuna tıklayın.',
        level2: 'Filtre menüsünde "İstanbul" değerini seçin veya yazıp Filtrele butonuna basın.',
        level3: 'C sütununu İstanbul olarak filtreleyin. Ankara ve İzmir satırları gizlenecektir.',
      },
      validate: (sheet) => validateTableFiltered(sheet, 'C', 1),
    },
  ],

  independentPractice: {
    goal: 'Depo Sevkiyat tablosunda Bölge sütununa (B sütunu) "Marmara" filtresi uygulayarak Ege sevkiyatını gizleyin.',
    scenario:
      'Depo müdürü Marmara bölgesine gidecek koli sayısını kontrol etmek istiyor. B sütununa Marmara filtresi uygulayın.',
    narrationText:
      'Bağımsız alıştırmada Bölge sütununa Marmara filtresi uygulayın.',
    workbookFactory: createFilterPracticeWorkbook,
    targetCell: 'B1',
    hints: [
      'Önce araç çubuğundaki Filtrele butonuna tıklayın.',
      'B sütun başlığındaki filtre menüsünden Marmara seçeneğini uygulayın.',
      'Ege satırı (2. satır) gizlendiğinde görev tamamlanacaktır.',
    ],
    validate: (sheet) => validateTableFiltered(sheet, 'B', 1),
  },

  quizQuestions: [
    {
      id: 'q_filter_1',
      question: 'Excel\'de bir sütuna filtre uygulandığında filtre kriterine uymayan satırlara ne olur?',
      options: [
        'Geçici olarak gizlenirler, filtre kaldırılınca geri gelirler',
        'Kalıcı olarak silinirler',
        'Yeni bir sayfaya taşınırlar',
        'Değerleri sıfırlanır',
      ],
      correctIndex: 0,
      explanation: 'Filtreleme işlemi asla satırları silmez; yalnızca görünürlüklerini geçici olarak kapatır.',
    },
    {
      id: 'q_filter_2',
      question: 'Tabloda filtrenin açık olduğunu en kolay nasıl anlarsınız?',
      options: [
        'Sütun başlıklarında küçük aşağı ok / filtre simgeleri görünür',
        'Ekran siyah renge dönüşür',
        'Hücreler kilitlenir ve yazı yazılamaz',
        'Excel otomatik olarak kapanır',
      ],
      correctIndex: 0,
      explanation: 'Filtre modu açıldığında tablo başlıklarının sağ tarafında aşağı ok / filtre simgeleri yer alır.',
    },
  ],

  initialWorkbookFactory: createFilterInitialWorkbook,
};
