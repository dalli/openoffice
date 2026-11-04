import type { zh } from './zh'

export const en = {
  splitView: 'Split view',
  splitSide: 'Side by side',
  splitStack: 'Top and bottom',
  splitUnsplit: 'Unsplit',
  splitUnavailable: 'Home, help and presentation views cannot be split.',
  splitSingleCopy: 'One document: opens a separate copy; changes do not sync.',
  splitLeftDocument: 'Left document',
  splitRightDocument: 'Right document',
  splitTopDocument: 'Top document',
  splitBottomDocument: 'Bottom document',
} satisfies Record<keyof typeof zh, string>
