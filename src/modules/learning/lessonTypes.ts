/**
 * Learning and Lesson state machine types for Excel Akademi AI
 * Stage 4: Universal Lesson Content Engine & Pedagogical System
 * Supports 8 distinct Excel lesson types with comprehensive pedagogical structures
 */

import type { Worksheet, Workbook, CellCoordinate, NumberFormat, HorizontalAlignment } from '../spreadsheet/types.ts';

/**
 * 8 Universal Excel lesson categories supported by the curriculum engine:
 */
export type LessonType =
  | 'concept'             // A. Kavram dersleri (Hücre, Çalışma Kitabı, Veri Türleri)
  | 'interface_action'    // B. Arayüz ve işlem dersleri (Şerit menü, Biçimlendirme, Kısayollar)
  | 'formula_function'   // C. Formül ve fonksiyon dersleri (TOPLA, EĞER, DÜŞEYARA)
  | 'data_management'     // D. Veri yönetimi dersleri (Sıralama, Filtreleme, Yinelenenler)
  | 'visualization'       // E. Görselleştirme dersleri (Grafikler, Mini Grafikler, Dashboard)
  | 'advanced_analysis'   // F. İleri analiz dersleri (PivotTable, Çapraz Tablolar)
  | 'automation'          // G. Otomasyon dersleri (Makrolar, Veri Temizleme)
  | 'project';            // H. Proje dersleri (Uçtan uca iş senaryoları)

/**
 * Compulsory 14-state lesson engine states:
 */
export type LessonEngineState =
  | 'lesson_intro'
  | 'explanation'
  | 'worked_example'
  | 'ready_for_task'
  | 'teacher_instruction'
  | 'waiting_for_student'
  | 'validating'
  | 'correct'
  | 'incorrect'
  | 'hint'
  | 'next_step'
  | 'independent_practice'
  | 'assessment'
  | 'completed';

export type LessonStepStatus =
  | 'explanation'
  | 'example'
  | 'instruction'
  | 'waiting_for_student'
  | 'validating'
  | 'success'
  | 'retry'
  | 'completed';

export type TaskActionType =
  | 'cell_selection'        // Hücre seçme
  | 'data_entry'            // Veri girme
  | 'formula_input'          // Formül yazma
  | 'formula_result'         // Formül sonucunu kontrol etme
  | 'cell_formatting'        // Hücre biçimlendirme
  | 'menu_action'            // Menüden işlem seçme
  | 'context_menu'           // Sağ tıklama menüsünü kullanma
  | 'keyboard_shortcut'      // Klavye kısayolu kullanma
  | 'insert_row_col'         // Satır veya sütun ekleme
  | 'data_sort'              // Veri sıralama
  | 'data_filter'            // Filtre uygulama
  | 'create_chart'           // Grafik oluşturma
  | 'pivottable_action'      // PivotTable işlemi
  | 'error_detection'        // Hata tespiti
  | 'auto_recalculation'     // Otomatik yeniden hesaplama
  | 'independent_challenge'; // Bağımsız problem çözme

export interface TaskValidationResult {
  isValid: boolean;
  message: string;
  hint?: string;
  expectedValue?: string | number;
  diagnosticDetail?: string;
}

export interface TaskHintLevels {
  level1: string; // Hafif ipucu: Dikkat edilmesi gereken nokta
  level2: string; // Yönlendirici ipucu: Nereye/hangi hücreye veya menüye bakılacağı
  level3: string; // Ayrıntılı yardım: Adım adım ne yapılması gerektiği
}

export interface TaskStep {
  id: string;
  taskNumber?: number;
  number?: number;
  title: string;
  actionType?: TaskActionType;
  instruction: string;
  hint?: string; // Default hint
  hints: TaskHintLevels; // 3-level graduated hint system
  narrationText: string;
  teacherIntroNarration?: string;
  targetCell?: string; // e.g. "B7"
  targetRange?: string; // e.g. "B2:B6"
  validate: (sheet: Worksheet, activeCell: CellCoordinate, isEditing?: boolean, selectionRange?: any) => TaskValidationResult;
}

export interface LessonSolvedExample {
  title: string;
  scenario: string;
  dataSnippet?: Array<{ label: string; value: string }>;
  formula?: string;
  result?: string;
  procedureSteps?: string[];
  explanation: string;
  steps: Array<string | { stepNumber?: number; instruction: string; tip?: string }>;
  audioNarration?: string;
}

export interface ExplanationTopic {
  id: string;
  number: number;
  title: string;
  content: string;
  keyTakeaway: string;
  audioNarration: string;
}

export interface LessonGlossaryItem {
  id: string;
  term: string;
  englishEquivalent?: string;
  definition: string;
  category?: string;
  example?: string;
}

export interface LessonUiLocation {
  elementName: string;
  ribbonTab: string; // 'Giriş' | 'Ekle' | 'Sayfa Düzeni' | 'Formüller' | 'Veri'
  group: string;     // 'Yazı Tipi' | 'Hizalama' | 'Sayı' | 'Stiller' | 'İşlev Kitaplığı'
  shortcut?: string; // e.g. 'Ctrl + B'
  description: string;
}

export interface LessonRealWorldScenario {
  id?: string;
  industry: string;       // e.g. 'Finans & Muhasebe', 'Perakende', 'Lojistik'
  role: string;           // e.g. 'Bütçe Uzmanı', 'Mağaza Müdürü'
  challenge: string;      // İş yerinde karşılaşılan somut problem
  solution: string;       // Excel aracı ile sağlanan çözüm
  dataDescription?: string;// Tablo yapısı
  appliedBenefit: string; // Zaman tasarrufu veya hata önleme kazanımı
}

export interface LessonMistakeItem {
  id?: string;
  title?: string;
  mistake: string;
  cause: string;
  solution: string;
  badExample?: string;
  goodExample?: string;
}

export interface LessonParameter {
  name: string;
  description: string;
  required: boolean;
  defaultValue?: string;
  example?: string;
}

export interface QuizQuestion {
  id: string;
  question: string;
  options: string[];
  correctIndex: number; // 0-indexed
  explanation: string;
  difficulty?: 'easy' | 'medium' | 'hard';
}

export interface IndependentPracticeSpec {
  title?: string;
  description?: string;
  scenario?: string;
  goal: string;
  targetCell: string;
  targetRange?: string;
  expectedFormula?: string[];
  expectedResult?: number | string;
  expectedFormat?: NumberFormat;
  expectedAlignment?: HorizontalAlignment;
  hint?: string;
  hints?: TaskHintLevels | string[];
  narrationText: string;
  workbookFactory: () => Workbook;
  validate?: (sheet: Worksheet) => TaskValidationResult;
  customValidator?: (sheet: Worksheet) => TaskValidationResult;
}

export interface AssessmentReport {
  lessonId: string;
  lessonTitle: string;
  completedAt: string;
  tasksCompleted: number;
  totalTasks: number;
  independentPracticeCompleted: boolean;
  totalWrongAttempts: number;
  totalHintsUsed: number;
  score: number; // 0 - 100
  masteredSkills: string[];
  topicsToReview: string[];
  nextRecommendedLesson: {
    id: string;
    title: string;
  };
}

export interface LessonCompletionCriteria {
  minGuidedTasksPercent: number; // e.g. 100
  requireIndependentPractice: boolean;
  minQuizScorePercent?: number; // e.g. 70
}

export interface LessonVisualizationSpec {
  chartTypes: Array<'column' | 'bar' | 'line' | 'pie' | 'area'>;
  dataRangeDescription: string;
  recommendedColors: string[];
}

export interface LessonDataManagementSpec {
  datasetRows: number;
  filterColumns: string[];
  sortKey: string;
  duplicateKeys?: string[];
}

export interface LessonAutomationSpec {
  macroName: string;
  macroTrigger: string;
  automatedActions: string[];
}

export interface LessonProjectSpec {
  businessDomain: string;
  deliverables: string[];
  phases: Array<{ name: string; taskIds: string[] }>;
}

export interface Lesson {
  id: string;
  number: number;
  title: string;
  lessonType: LessonType;
  category: string;
  level: 'Başlangıç' | 'Orta' | 'İleri' | string;
  estimatedMinutes: number;
  description: string;
  learningObjective: string;
  learningOutcomes?: string[];
  prerequisites?: Array<{ id: string; title: string }>;
  introNarration: string;
  topics: ExplanationTopic[];
  glossary?: LessonGlossaryItem[];
  uiLocations?: LessonUiLocation[];
  realWorldScenarios?: LessonRealWorldScenario[];
  commonMistakes?: LessonMistakeItem[];
  bestPractices?: string[];
  explanationContent?: {
    overview: string;
    syntax?: string;
    parameters?: LessonParameter[];
    procedureSteps?: string[];
    bestPractices: string[];
    commonMistakes: string[];
  };
  solvedExample: LessonSolvedExample;
  alternativeExamples?: LessonSolvedExample[];
  tasks: TaskStep[];
  independentPractice: IndependentPracticeSpec;
  quizQuestions?: QuizQuestion[];
  summary?: string[];
  completionCriteria?: LessonCompletionCriteria;
  visualizationSpec?: LessonVisualizationSpec;
  dataManagementSpec?: LessonDataManagementSpec;
  automationSpec?: LessonAutomationSpec;
  projectSpec?: LessonProjectSpec;
  initialWorkbookFactory: () => Workbook;
}

export interface LessonTypeDescriptor {
  type: LessonType;
  titleTr: string;
  categoryName: string;
  description: string;
  badgeBg: string;
  badgeText: string;
  iconName: string;
  focusAreas: string[];
}

export const LESSON_TYPE_DESCRIPTORS: Record<LessonType, LessonTypeDescriptor> = {
  concept: {
    type: 'concept',
    titleTr: 'Kavram Dersi',
    categoryName: 'Temel Kavramlar',
    description: 'Excel çalışma alanı, hücreler, sayfalar ve temel veri modelleri.',
    badgeBg: 'bg-blue-50 border-blue-200',
    badgeText: 'text-blue-800',
    iconName: 'BookOpen',
    focusAreas: ['Çalışma Kitabı Mantığı', 'Hücre Koordinatları', 'Veri Tipleri'],
  },
  interface_action: {
    type: 'interface_action',
    titleTr: 'Arayüz & İşlem Dersi',
    categoryName: 'Arayüz ve Biçimlendirme',
    description: 'Şerit menü, hücre stilleri, para birimi, hizalama ve kısayollar.',
    badgeBg: 'bg-emerald-50 border-emerald-200',
    badgeText: 'text-emerald-800',
    iconName: 'Layers',
    focusAreas: ['Yazı Tipleri & Renkler', 'Sayı Biçimlendirme (₺, %)', 'Klavye Kısayolları'],
  },
  formula_function: {
    type: 'formula_function',
    titleTr: 'Formül & Fonksiyon Dersi',
    categoryName: 'Formüller ve Hesaplama',
    description: 'Matematiksel ve mantıksal Excel fonksiyonları, sözdizimi ve dinamik hesaplama.',
    badgeBg: 'bg-purple-50 border-purple-200',
    badgeText: 'text-purple-800',
    iconName: 'Code2',
    focusAreas: ['Sözdizimi (Syntax)', 'Parametre Kuralları', 'Hata Denetimi (#DEĞER!, #AD?)'],
  },
  data_management: {
    type: 'data_management',
    titleTr: 'Veri Yönetimi Dersi',
    categoryName: 'Veri ve Tablo Yönetimi',
    description: 'Büyük listeleri sıralama, çoklu filtreleme, yinelenenleri kaldırma ve tablolaştırma.',
    badgeBg: 'bg-amber-50 border-amber-200',
    badgeText: 'text-amber-800',
    iconName: 'Database',
    focusAreas: ['A-Z / Z-A Sıralama', 'Özel Filtreler', 'Veri Doğrulama'],
  },
  visualization: {
    type: 'visualization',
    titleTr: 'Görselleştirme Dersi',
    categoryName: 'Grafik ve Sunum',
    description: 'Sütun, çizgi ve pasta grafikler ile verileri görsel dashboardlara dönüştürme.',
    badgeBg: 'bg-rose-50 border-rose-200',
    badgeText: 'text-rose-800',
    iconName: 'BarChart3',
    focusAreas: ['Grafik Türü Seçimi', 'Veri Serisi Tanımlama', 'Etiket ve Eksen Tasarımı'],
  },
  advanced_analysis: {
    type: 'advanced_analysis',
    titleTr: 'İleri Analiz Dersi',
    categoryName: 'PivotTable & Analitik',
    description: 'Özet Tablolar (PivotTable), dilimleyiciler ve çok boyutlu çapraz analizler.',
    badgeBg: 'bg-indigo-50 border-indigo-200',
    badgeText: 'text-indigo-800',
    iconName: 'TrendingUp',
    focusAreas: ['Satır/Sütun Boyutları', 'Özet Değer Alanları', 'Yüzdelik Dağılımlar'],
  },
  automation: {
    type: 'automation',
    titleTr: 'Otomasyon Dersi',
    categoryName: 'Makro & Otomasyon',
    description: 'Tekrarlayan iş süreçlerini tek tuşla otomatikleştiren makro akışları.',
    badgeBg: 'bg-teal-50 border-teal-200',
    badgeText: 'text-teal-800',
    iconName: 'Zap',
    focusAreas: ['Makro Kaydı', 'Görev Tetikleyicileri', 'Toplu Veri İşleme'],
  },
  project: {
    type: 'project',
    titleTr: 'Uçtan Uca Proje Dersi',
    categoryName: 'İş Hayatı Projeleri',
    description: 'Birden fazla Excel becerisini bir araya getiren gerçek bir şirket senaryosu çözümü.',
    badgeBg: 'bg-orange-50 border-orange-200',
    badgeText: 'text-orange-800',
    iconName: 'Briefcase',
    focusAreas: ['Çok Sayfalı Raporlama', 'Dinamik Modeller', 'Yönetim Özeti'],
  },
};

export function getLessonTypeDescriptor(type: LessonType): LessonTypeDescriptor {
  return LESSON_TYPE_DESCRIPTORS[type] || LESSON_TYPE_DESCRIPTORS.concept;
}

export interface UserLessonProgress {
  lessonId: string;
  completed: boolean;
  currentTaskIndex: number;
  completedTasks: string[];
  lastStudiedAt: string;
  score?: number;
  quizScore?: number;
  engineState?: LessonEngineState;
  hintsUsedCount?: number;
  wrongAttemptsCount?: number;
  independentCompleted?: boolean;
}

export interface OverallLearningProgress {
  totalCompletedLessons: number;
  totalTimeSpentMinutes: number;
  dailyGoalMinutes: number;
  dailyMinutesCompleted: number;
  lastLessonId: string;
  recommendedLessonId: string;
  lessonProgress: Record<string, UserLessonProgress>;
}
