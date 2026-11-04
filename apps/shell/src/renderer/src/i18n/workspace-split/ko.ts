import type { zh } from './zh'

export const ko = {
  splitView: '화면 나누기',
  splitSide: '좌우로 보기',
  splitStack: '위아래로 보기',
  splitUnsplit: '나누기 해제',
  splitUnavailable: '홈·도움말·발표 화면은 나눌 수 없습니다.',
  splitSingleCopy: '문서가 하나입니다. 별도 사본을 열며 변경은 동기화되지 않습니다.',
  splitLeftDocument: '왼쪽 문서',
  splitRightDocument: '오른쪽 문서',
  splitTopDocument: '위쪽 문서',
  splitBottomDocument: '아래쪽 문서',
} satisfies Record<keyof typeof zh, string>
