import type { zh } from './zh'

export const zhTW = {
  splitView: '分割畫面',
  splitSide: '左右並排',
  splitStack: '上下排列',
  splitUnsplit: '取消分割',
  splitUnavailable: '首頁、說明和簡報畫面無法分割。',
  splitSingleCopy: '僅有一份文件：將開啟獨立副本，變更不會同步。',
  splitLeftDocument: '左側文件',
  splitRightDocument: '右側文件',
  splitTopDocument: '上方文件',
  splitBottomDocument: '下方文件',
} satisfies Record<keyof typeof zh, string>
