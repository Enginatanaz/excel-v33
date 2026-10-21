/**
 * Curriculum Lesson Catalog
 * Stage 5: Comprehensive Universal Lesson Registry
 * Covering Concepts, Interfaces, Formulas (SUM, IF, COUNTIF, SUMIFS, VLOOKUP, XLOOKUP),
 * Data Management (Sorting, Filtering), and Visualization (Charts).
 */

import { excelIntroductionLesson } from './excelIntroduction.ts';
import { sumFunctionLesson } from './sumFunction.ts';
import { cellFormattingLesson } from './cellFormatting.ts';
import { ifFunctionLesson } from './ifFunction.ts';
import { countifFunctionLesson } from './countifFunction.ts';
import { sumifsFunctionLesson } from './sumifsFunction.ts';
import { vlookupFunctionLesson } from './vlookupFunction.ts';
import { xlookupFunctionLesson } from './xlookupFunction.ts';
import { dataSortLesson } from './dataSortLesson.ts';
import { dataFilterLesson } from './dataFilterLesson.ts';
import { chartCreationLesson } from './chartCreationLesson.ts';
import type { Lesson } from '../../modules/learning/lessonTypes.ts';

export const allLessons: Lesson[] = [
  excelIntroductionLesson,
  sumFunctionLesson,
  cellFormattingLesson,
  ifFunctionLesson,
  countifFunctionLesson,
  sumifsFunctionLesson,
  vlookupFunctionLesson,
  xlookupFunctionLesson,
  dataSortLesson,
  dataFilterLesson,
  chartCreationLesson,
];

export function getLessonById(id: string): Lesson | undefined {
  return allLessons.find((l) => l.id === id);
}

export {
  excelIntroductionLesson,
  sumFunctionLesson,
  cellFormattingLesson,
  ifFunctionLesson,
  countifFunctionLesson,
  sumifsFunctionLesson,
  vlookupFunctionLesson,
  xlookupFunctionLesson,
  dataSortLesson,
  dataFilterLesson,
  chartCreationLesson,
};
