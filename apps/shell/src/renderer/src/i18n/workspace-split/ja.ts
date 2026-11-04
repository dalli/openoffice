import type { zh } from './zh'

export const ja = {
  splitView: '画面分割',
  splitSide: '左右に表示',
  splitStack: '上下に表示',
  splitUnsplit: '分割を解除',
  splitUnavailable: 'ホーム・ヘルプ・発表画面は分割できません。',
  splitSingleCopy: '文書は1つです。別のコピーを開き、変更は同期しません。',
  splitLeftDocument: '左側の文書',
  splitRightDocument: '右側の文書',
  splitTopDocument: '上側の文書',
  splitBottomDocument: '下側の文書',
} satisfies Record<keyof typeof zh, string>
