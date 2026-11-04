import type { zh } from './zh'

export const th = {
  splitView: 'แบ่งมุมมอง',
  splitSide: 'ดูซ้ายและขวา',
  splitStack: 'ดูบนและล่าง',
  splitUnsplit: 'ยกเลิกการแบ่ง',
  splitUnavailable: 'ไม่สามารถแบ่งหน้าหลัก หน้าช่วยเหลือ และหน้าการนำเสนอได้',
  splitSingleCopy: 'มีเอกสารเดียว: จะเปิดสำเนาแยกต่างหาก โดยไม่ซิงค์การเปลี่ยนแปลง',
  splitLeftDocument: 'เอกสารด้านซ้าย',
  splitRightDocument: 'เอกสารด้านขวา',
  splitTopDocument: 'เอกสารด้านบน',
  splitBottomDocument: 'เอกสารด้านล่าง',
} satisfies Record<keyof typeof zh, string>
