import type { zh } from './zh'

export const de = {
  splitView: 'Ansicht teilen',
  splitSide: 'Nebeneinander',
  splitStack: 'Untereinander',
  splitUnsplit: 'Teilung aufheben',
  splitUnavailable: 'Startseite, Hilfe und Präsentation lassen sich nicht teilen.',
  splitSingleCopy:
    'Ein Dokument: Eine separate Kopie wird geöffnet; Änderungen werden nicht synchronisiert.',
  splitLeftDocument: 'Linkes Dokument',
  splitRightDocument: 'Rechtes Dokument',
  splitTopDocument: 'Oberes Dokument',
  splitBottomDocument: 'Unteres Dokument',
} satisfies Record<keyof typeof zh, string>
