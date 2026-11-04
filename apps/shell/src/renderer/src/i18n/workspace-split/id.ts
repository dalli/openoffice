import type { zh } from './zh'

export const id = {
  splitView: 'Bagi tampilan',
  splitSide: 'Berdampingan',
  splitStack: 'Atas dan bawah',
  splitUnsplit: 'Batalkan pembagian',
  splitUnavailable: 'Tampilan Beranda, bantuan, dan presentasi tidak dapat dibagi.',
  splitSingleCopy: 'Satu dokumen: membuka salinan terpisah; perubahan tidak disinkronkan.',
  splitLeftDocument: 'Dokumen kiri',
  splitRightDocument: 'Dokumen kanan',
  splitTopDocument: 'Dokumen atas',
  splitBottomDocument: 'Dokumen bawah',
} satisfies Record<keyof typeof zh, string>
