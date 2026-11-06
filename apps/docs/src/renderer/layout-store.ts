import type { SetStateAction } from 'react'
import type { FileActionContext } from './file-actions'
import { EMPTY_HF_VARIANTS } from './doc-state'

/** The pending layout and style values already consumed by the document save path. */
export type DocsLayoutState = Pick<
  FileActionContext,
  | 'section'
  | 'sections'
  | 'sectionDirty'
  | 'sectionsDirty'
  | 'trailingStartType'
  | 'pageColor'
  | 'pageColorDirty'
  | 'header'
  | 'headerDirty'
  | 'footer'
  | 'footerDirty'
  | 'hfVariants'
  | 'hfVariantsDirty'
  | 'sectionHfEdits'
  | 'hfLinks'
  | 'titlePg'
  | 'titlePgDirty'
  | 'evenOddHf'
  | 'evenOddHfDirty'
  | 'mirrorMargins'
  | 'mirrorMarginsDirty'
  | 'pgNumEdit'
  | 'pgNumDirtySections'
  | 'defaultFonts'
  | 'styleUpserts'
  | 'watermark'
  | 'watermarkDirty'
  | 'watermarkStyle'
  | 'watermarkPicture'
>

const CONTENT_KEYS = [
  'section',
  'sections',
  'trailingStartType',
  'pageColor',
  'header',
  'footer',
  'hfVariants',
  'sectionHfEdits',
  'hfLinks',
  'titlePg',
  'evenOddHf',
  'mirrorMargins',
  'pgNumEdit',
  'defaultFonts',
  'styleUpserts',
  'watermark',
  'watermarkStyle',
  'watermarkPicture',
] as const

/** App-owned immutable state, observed synchronously by editor tools and React. */
export class DocsLayoutStore {
  private state: DocsLayoutState = {
    section: null,
    sections: [],
    sectionDirty: false,
    sectionsDirty: [],
    trailingStartType: null,
    pageColor: null,
    pageColorDirty: false,
    header: null,
    headerDirty: false,
    footer: null,
    footerDirty: false,
    hfVariants: EMPTY_HF_VARIANTS,
    hfVariantsDirty: [],
    sectionHfEdits: {},
    hfLinks: {},
    titlePg: false,
    titlePgDirty: false,
    evenOddHf: false,
    evenOddHfDirty: false,
    mirrorMargins: false,
    mirrorMarginsDirty: false,
    pgNumEdit: null,
    pgNumDirtySections: [],
    defaultFonts: undefined,
    styleUpserts: {},
    watermark: null,
    watermarkDirty: false,
    watermarkStyle: null,
    watermarkPicture: null,
  }
  private listeners = new Set<() => void>()
  private contentRevision = 0

  get revision(): number {
    return this.contentRevision
  }
  getSnapshot = (): DocsLayoutState => this.state
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  private update(patch: Partial<DocsLayoutState>): void {
    const previous = this.state
    const next = { ...previous, ...patch }
    if (
      Object.keys(patch).every((key) =>
        Object.is(previous[key as keyof DocsLayoutState], next[key as keyof DocsLayoutState]),
      )
    )
      return
    if (CONTENT_KEYS.some((key) => !Object.is(previous[key], next[key]))) this.contentRevision++
    this.state = next
    for (const listener of this.listeners) listener()
  }

  private set<K extends keyof DocsLayoutState>(
    key: K,
    action: SetStateAction<DocsLayoutState[K]>,
  ): void {
    this.update({ [key]: typeof action === 'function' ? action(this.state[key]) : action })
  }
  setSection = (action: SetStateAction<DocsLayoutState['section']>): void =>
    this.set('section', action)
  setSections = (action: SetStateAction<DocsLayoutState['sections']>): void =>
    this.set('sections', action)
  setSectionDirty = (action: SetStateAction<DocsLayoutState['sectionDirty']>): void =>
    this.set('sectionDirty', action)
  setSectionsDirty = (action: SetStateAction<DocsLayoutState['sectionsDirty']>): void =>
    this.set('sectionsDirty', action)
  setTrailingStartType = (action: SetStateAction<DocsLayoutState['trailingStartType']>): void =>
    this.set('trailingStartType', action)
  setPageColor = (action: SetStateAction<DocsLayoutState['pageColor']>): void =>
    this.set('pageColor', action)
  setPageColorDirty = (action: SetStateAction<DocsLayoutState['pageColorDirty']>): void =>
    this.set('pageColorDirty', action)
  setHeader = (action: SetStateAction<DocsLayoutState['header']>): void =>
    this.set('header', action)
  setHeaderDirty = (action: SetStateAction<DocsLayoutState['headerDirty']>): void =>
    this.set('headerDirty', action)
  setFooter = (action: SetStateAction<DocsLayoutState['footer']>): void =>
    this.set('footer', action)
  setFooterDirty = (action: SetStateAction<DocsLayoutState['footerDirty']>): void =>
    this.set('footerDirty', action)
  setHfVariants = (action: SetStateAction<DocsLayoutState['hfVariants']>): void =>
    this.set('hfVariants', action)
  setHfVariantsDirty = (action: SetStateAction<DocsLayoutState['hfVariantsDirty']>): void =>
    this.set('hfVariantsDirty', action)
  setSectionHfEdits = (action: SetStateAction<DocsLayoutState['sectionHfEdits']>): void =>
    this.set('sectionHfEdits', action)
  setHfLinks = (action: SetStateAction<DocsLayoutState['hfLinks']>): void =>
    this.set('hfLinks', action)
  setTitlePg = (action: SetStateAction<DocsLayoutState['titlePg']>): void =>
    this.set('titlePg', action)
  setTitlePgDirty = (action: SetStateAction<DocsLayoutState['titlePgDirty']>): void =>
    this.set('titlePgDirty', action)
  setEvenOddHf = (action: SetStateAction<DocsLayoutState['evenOddHf']>): void =>
    this.set('evenOddHf', action)
  setEvenOddHfDirty = (action: SetStateAction<DocsLayoutState['evenOddHfDirty']>): void =>
    this.set('evenOddHfDirty', action)
  setMirrorMargins = (action: SetStateAction<DocsLayoutState['mirrorMargins']>): void =>
    this.set('mirrorMargins', action)
  setMirrorMarginsDirty = (action: SetStateAction<DocsLayoutState['mirrorMarginsDirty']>): void =>
    this.set('mirrorMarginsDirty', action)
  setPgNumEdit = (action: SetStateAction<DocsLayoutState['pgNumEdit']>): void =>
    this.set('pgNumEdit', action)
  setPgNumDirtySections = (action: SetStateAction<DocsLayoutState['pgNumDirtySections']>): void =>
    this.set('pgNumDirtySections', action)
  setDefaultFonts = (action: SetStateAction<DocsLayoutState['defaultFonts']>): void =>
    this.set('defaultFonts', action)
  setStyleUpserts = (action: SetStateAction<DocsLayoutState['styleUpserts']>): void =>
    this.set('styleUpserts', action)
  setWatermark = (action: SetStateAction<DocsLayoutState['watermark']>): void =>
    this.set('watermark', action)
  setWatermarkDirty = (action: SetStateAction<DocsLayoutState['watermarkDirty']>): void =>
    this.set('watermarkDirty', action)
  setWatermarkStyle = (action: SetStateAction<DocsLayoutState['watermarkStyle']>): void =>
    this.set('watermarkStyle', action)
  setWatermarkPicture = (action: SetStateAction<DocsLayoutState['watermarkPicture']>): void =>
    this.set('watermarkPicture', action)

  /** Retained only in this renderer, alongside the body and review snapshots. */
  capture(): DocsLayoutState {
    return { ...this.state }
  }

  restore(snapshot: DocsLayoutState): void {
    // The adapter rejects a changed parsed baseline before restoring. Keeping
    // the exact pending flags avoids creating parts absent from that baseline.
    this.update(snapshot)
  }
}
