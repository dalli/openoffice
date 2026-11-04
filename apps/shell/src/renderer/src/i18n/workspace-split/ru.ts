import type { zh } from './zh'

export const ru = {
  splitView: 'Разделить окно',
  splitSide: 'Рядом',
  splitStack: 'Сверху и снизу',
  splitUnsplit: 'Убрать разделение',
  splitUnavailable: 'Главную страницу, справку и презентацию разделить нельзя.',
  splitSingleCopy: 'Один документ: откроется отдельная копия; изменения не синхронизируются.',
  splitLeftDocument: 'Документ слева',
  splitRightDocument: 'Документ справа',
  splitTopDocument: 'Документ сверху',
  splitBottomDocument: 'Документ снизу',
} satisfies Record<keyof typeof zh, string>
