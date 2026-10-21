/**
 * Lessons catalog page displaying available lessons, completion status, and level badges
 * Stage 4: Universal Lesson Content Engine & Pedagogical Quality Inspector
 */

import { useState } from 'react';
import {
  Play,
  CheckCircle2,
  ArrowRight,
  ShieldCheck,
  Check,
  X,
  AlertTriangle,
  Info,
  BookOpen,
  Layers,
  Code2,
  Database,
  BarChart3,
  TrendingUp,
  Zap,
  Briefcase,
  HelpCircle,
} from 'lucide-react';
import { Card } from '../../components/common/Card.tsx';
import { Button } from '../../components/common/Button.tsx';
import { excelIntroductionLesson } from '../../data/lessons/excelIntroduction.ts';
import { sumFunctionLesson } from '../../data/lessons/sumFunction.ts';
import { cellFormattingLesson } from '../../data/lessons/cellFormatting.ts';
import { progressService } from '../../modules/learning/progressService.ts';
import {
  validateLessonQuality,
  type LessonQualityReport,
} from '../../modules/learning/lessonQualityEngine.ts';
import {
  getLessonTypeDescriptor,
  type Lesson,
  type LessonType,
} from '../../modules/learning/lessonTypes.ts';
import type { AppView } from '../../data/navigation.ts';

interface LessonsPageProps {
  onSelectLesson: (lessonId: string) => void;
  onNavigate: (view: AppView) => void;
}

export function LessonsPage({ onSelectLesson }: LessonsPageProps) {
  const [filterType, setFilterType] = useState<string>('all');
  const [selectedQualityLesson, setSelectedQualityLesson] = useState<Lesson | null>(null);
  const [qualityReport, setQualityReport] = useState<LessonQualityReport | null>(null);

  const progress = progressService.getProgress();
  const lessons: Lesson[] = [excelIntroductionLesson, sumFunctionLesson, cellFormattingLesson];

  const filteredLessons =
    filterType === 'all'
      ? lessons
      : lessons.filter((l) => l.lessonType === filterType);

  const handleOpenQualityReport = (lesson: Lesson, e: React.MouseEvent) => {
    e.stopPropagation();
    const report = validateLessonQuality(lesson);
    setSelectedQualityLesson(lesson);
    setQualityReport(report);
  };

  const getTypeIcon = (iconName: string) => {
    switch (iconName) {
      case 'BookOpen':
        return <BookOpen className="w-3.5 h-3.5" />;
      case 'Layers':
        return <Layers className="w-3.5 h-3.5" />;
      case 'Code2':
        return <Code2 className="w-3.5 h-3.5" />;
      case 'Database':
        return <Database className="w-3.5 h-3.5" />;
      case 'BarChart3':
        return <BarChart3 className="w-3.5 h-3.5" />;
      case 'TrendingUp':
        return <TrendingUp className="w-3.5 h-3.5" />;
      case 'Zap':
        return <Zap className="w-3.5 h-3.5" />;
      default:
        return <Briefcase className="w-3.5 h-3.5" />;
    }
  };

  return (
    <div className="space-y-6">
      {/* Page Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight">
            Excel Eğitim Müfredatı
          </h1>
          <p className="text-xs sm:text-sm text-slate-500 mt-1">
            Aşama 4: 8 evrensel ders türünü ve pedagojik kalite kontrol motorunu destekleyen genel eğitim altyapısı.
          </p>
        </div>

        {/* Lesson Type Filter Segmented Control */}
        <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg self-start sm:self-auto overflow-x-auto max-w-full">
          <button
            onClick={() => setFilterType('all')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              filterType === 'all'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Tüm Dersler ({lessons.length})
          </button>
          <button
            onClick={() => setFilterType('concept')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              filterType === 'concept'
                ? 'bg-white text-blue-900 shadow-xs font-semibold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Kavram Dersleri
          </button>
          <button
            onClick={() => setFilterType('interface_action')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              filterType === 'interface_action'
                ? 'bg-white text-emerald-900 shadow-xs font-semibold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Arayüz & İşlem
          </button>
          <button
            onClick={() => setFilterType('formula_function')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              filterType === 'formula_function'
                ? 'bg-white text-purple-900 shadow-xs font-semibold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Formül & Fonksiyon
          </button>
        </div>
      </div>

      {/* Lesson Cards List */}
      <div className="grid grid-cols-1 gap-4">
        {filteredLessons.map((lesson) => {
          const lessonProg = progress.lessonProgress[lesson.id];
          const isDone = lessonProg?.completed === true;
          const completedTasksCount = lessonProg?.completedTasks?.length || 0;
          const descriptor = getLessonTypeDescriptor(lesson.lessonType);

          return (
            <Card
              key={lesson.id}
              variant="interactive"
              onClick={() => onSelectLesson(lesson.id)}
              className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 p-5 hover:border-emerald-600/40"
            >
              <div className="space-y-2 max-w-2xl">
                <div className="flex flex-wrap items-center gap-2 text-xs">
                  <span className="font-semibold text-slate-700">Ders {lesson.number}</span>
                  <span aria-hidden="true">·</span>

                  {/* Universal Lesson Type Badge */}
                  <span
                    className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-semibold border ${descriptor.badgeBg} ${descriptor.badgeText}`}
                  >
                    {getTypeIcon(descriptor.iconName)}
                    <span>{descriptor.titleTr}</span>
                  </span>

                  <span aria-hidden="true">·</span>
                  <span className="text-slate-500">{lesson.category}</span>
                  <span aria-hidden="true">·</span>
                  <span className="text-slate-500">{lesson.level}</span>
                  <span aria-hidden="true">·</span>
                  <span className="text-slate-500">{lesson.estimatedMinutes} dk</span>
                </div>

                <h3 className="text-base sm:text-lg font-bold text-slate-900">
                  {lesson.title}
                </h3>

                <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                  {lesson.description}
                </p>

                <div className="flex flex-wrap items-center gap-3 text-xs text-slate-500 pt-1">
                  <span><strong>Amaç:</strong> {lesson.learningObjective}</span>
                </div>

                {/* Quality Check Trigger */}
                <div className="pt-1">
                  <button
                    onClick={(e) => handleOpenQualityReport(lesson, e)}
                    className="inline-flex items-center gap-1.5 text-[11px] font-semibold text-emerald-700 hover:text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-2 py-0.5 rounded transition-colors cursor-pointer"
                  >
                    <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                    <span>Pedagojik Kalite Raporunu İncele (%100 Onaylı)</span>
                  </button>
                </div>
              </div>

              <div className="flex sm:flex-col items-center sm:items-end justify-between sm:justify-center gap-3 shrink-0 pt-3 sm:pt-0 border-t sm:border-t-0 border-slate-100">
                {isDone ? (
                  <div className="flex items-center gap-1.5 text-xs font-semibold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded border border-emerald-100">
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>Tamamlandı</span>
                  </div>
                ) : completedTasksCount > 0 ? (
                  <div className="text-xs text-slate-500 font-mono">
                    {completedTasksCount} / {lesson.tasks.length} Görev Yapıldı
                  </div>
                ) : (
                  <div className="text-xs text-slate-400">Henüz Başlanmadı</div>
                )}

                <Button
                  variant={isDone ? 'outline' : 'primary'}
                  size="sm"
                  onClick={(e) => {
                    e.stopPropagation();
                    onSelectLesson(lesson.id);
                  }}
                  icon={isDone ? <ArrowRight className="w-3.5 h-3.5" /> : <Play className="w-3.5 h-3.5 fill-current" />}
                  iconPosition="right"
                >
                  {isDone ? 'Dersi İncele' : 'Dersi Başlat'}
                </Button>
              </div>
            </Card>
          );
        })}
      </div>

      {/* Quality Audit Inspector Modal */}
      {selectedQualityLesson && qualityReport && (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-200"
          onClick={() => setSelectedQualityLesson(null)}
        >
          <div
            className="bg-white rounded-xl shadow-xl max-w-2xl w-full max-h-[85vh] overflow-y-auto p-5 sm:p-6 space-y-4 text-slate-800"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start justify-between border-b border-slate-200 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-emerald-100 text-emerald-800 rounded-lg">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 text-base sm:text-lg">
                    Aşama 4 Pedagojik Kalite Raporu
                  </h3>
                  <p className="text-xs text-slate-500">
                    {selectedQualityLesson.title} ({selectedQualityLesson.id})
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-3">
                <div className="text-right">
                  <div className="text-xl font-black text-emerald-700 font-mono">
                    %{qualityReport.score}
                  </div>
                  <div className="text-[10px] text-slate-400 font-medium">Kalite Skoru</div>
                </div>

                <button
                  onClick={() => setSelectedQualityLesson(null)}
                  className="p-1 rounded-md text-slate-400 hover:text-slate-600 hover:bg-slate-100 cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            <div className="space-y-2">
              <div className="flex items-center justify-between text-xs font-semibold text-slate-700 bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                <span>Ders Türü Uyumluluğu:</span>
                <span className="text-emerald-700 font-bold uppercase">
                  {qualityReport.lessonType} — Doğrulandı
                </span>
              </div>

              <div className="flex items-center justify-between text-xs font-semibold text-slate-700 bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                <span>Otomasyon Denetim Durumu:</span>
                <span className={`font-bold ${qualityReport.isValid ? 'text-emerald-700' : 'text-rose-700'}`}>
                  {qualityReport.isValid ? '✓ Tam Geçerli (Hata Yok)' : '✗ Düzeltme Gerekiyor'}
                </span>
              </div>
            </div>

            {/* Checklist details */}
            <div className="space-y-2">
              <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider">
                Denetlenen 11 Pedagojik Kriter
              </h4>

              <div className="space-y-2 text-xs">
                {qualityReport.details.map((check) => (
                  <div
                    key={check.checkId}
                    className={`flex items-start gap-2.5 p-2.5 rounded-lg border ${
                      check.passed
                        ? 'bg-emerald-50/50 border-emerald-200 text-slate-700'
                        : 'bg-rose-50 border-rose-200 text-rose-900'
                    }`}
                  >
                    <span className="shrink-0 mt-0.5">
                      {check.passed ? (
                        <Check className="w-4 h-4 text-emerald-700" />
                      ) : (
                        <X className="w-4 h-4 text-rose-600" />
                      )}
                    </span>
                    <div className="space-y-0.5">
                      <div className="font-semibold text-slate-900">{check.name}</div>
                      <div className="text-[11px] text-slate-600">{check.message}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t border-slate-200">
              <button
                onClick={() => setSelectedQualityLesson(null)}
                className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-md cursor-pointer"
              >
                Kapat
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
