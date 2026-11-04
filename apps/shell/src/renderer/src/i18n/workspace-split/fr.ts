import type { zh } from './zh'

export const fr = {
  splitView: 'Diviser la vue',
  splitSide: 'Côte à côte',
  splitStack: 'Haut et bas',
  splitUnsplit: 'Annuler la division',
  splitUnavailable: 'Les vues accueil, aide et présentation ne peuvent pas être divisées.',
  splitSingleCopy:
    'Un document : une copie distincte sera ouverte ; les modifications ne sont pas synchronisées.',
  splitLeftDocument: 'Document de gauche',
  splitRightDocument: 'Document de droite',
  splitTopDocument: 'Document du haut',
  splitBottomDocument: 'Document du bas',
} satisfies Record<keyof typeof zh, string>
