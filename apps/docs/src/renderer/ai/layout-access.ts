import type { SectionInfo } from '@genoffice/docx-engine'
import {
  applySectionSettings,
  applyTitlePg,
  applyPageNumType,
  applySectionStartType,
  pendingHeadingLevel,
} from '@genoffice/docx-engine'
import type { FileActionContext } from '../file-actions'
import type { DocsLayoutStore, DocsLayoutState } from '../layout-store'
import { resolveHf, hfCommitTarget, type HfSectionState } from '../hf-sections'
import { liveSections } from '../pagination-sections'
import type { BlockBox } from '../pagination'
import { applyHfText, hfEditText } from '../editor/hf-text'
import { applyResolvedPageSetup, describeSection, type AiPageSetupAccess } from './page-setup'
import { patchPendingSectPr, sectionIndexAtBlock } from './pending-sections'
import type { AiStyleInfo } from './style-ops'
import type { AiDocExtras, AiHeaderFooterAccess } from './tools'

/** A section's XML with the same pending values the existing save path writes. */
export function effectiveLayoutSectPr(state: DocsLayoutState, idx: number): string | null {
  const sec = state.sections[idx]
  if (!sec) return null
  if (sec.pendingBreak) return sec.sectPrXml
  const isLast = idx === state.sections.length - 1
  let xml = sec.sectPrXml
  if (isLast ? state.sectionDirty : state.sectionsDirty.includes(idx)) {
    xml = applySectionSettings(xml, (isLast && state.section) || sec.settings)
    xml = applyTitlePg(xml, isLast ? state.titlePg : sec.titlePg)
  } else if (isLast && state.titlePgDirty) xml = applyTitlePg(xml, state.titlePg)
  if (isLast ? state.pgNumEdit !== null : state.pgNumDirtySections.includes(idx)) {
    const pg = isLast ? state.pgNumEdit : { fmt: sec.pageNumberFmt, start: sec.pageNumberStart }
    xml = applyPageNumType(xml, pg?.fmt, pg?.start)
  }
  return xml
}

export interface DocsLayoutAccessDeps {
  store: DocsLayoutStore
  getContext(): FileActionContext
  isLocked(): boolean
  getDeletedBreaks(): Set<number> | undefined
  insertBreak(
    type: SectionInfo['startType'],
    afterBlockIndex: number,
    copy: { sectPr: string; sections: SectionInfo[]; deferSave: boolean },
  ): void
}

/** Both AI hosts use the existing section/header/style model and pending save stores. */
export function createDocsLayoutAccess(deps: DocsLayoutAccessDeps): {
  hf: AiHeaderFooterAccess
  extras: AiDocExtras
  pageSetup(deferSave?: boolean): AiPageSetupAccess
} {
  const { store } = deps
  const hfState = (): HfSectionState => {
    const state = store.getSnapshot()
    const ctx = deps.getContext()
    const present: Array<{ docxIndex: number }> = []
    ctx.editor?.state.doc.forEach((node) => {
      if (typeof node.attrs.docxIndex === 'number')
        present.push({ docxIndex: node.attrs.docxIndex })
    })
    return {
      sections: liveSections(
        state.sections,
        present as BlockBox[],
        undefined,
        deps.getDeletedBreaks(),
      ),
      edits: state.sectionHfEdits,
      links: state.hfLinks,
      docStrips: { header: state.header, footer: state.footer, variants: state.hfVariants },
      parsed: ctx.doc?.parsed ?? null,
    }
  }
  const hf: AiHeaderFooterAccess = {
    read() {
      const state = store.getSnapshot()
      const live = hfState()
      const textOf = (kind: 'header' | 'footer', variant: 'default' | 'first' | 'even') => {
        const value = resolveHf(
          live,
          kind === 'header' ? 0 : Math.max(0, live.sections.length - 1),
          kind,
          variant,
        ).value
        return value ? hfEditText(value) : ''
      }
      return {
        header: textOf('header', 'default'),
        footer: textOf('footer', 'default'),
        headerFirst: state.titlePg ? textOf('header', 'first') : null,
        footerFirst: state.titlePg ? textOf('footer', 'first') : null,
        headerEven: state.evenOddHf ? textOf('header', 'even') : null,
        footerEven: state.evenOddHf ? textOf('footer', 'even') : null,
        titlePg: state.titlePg,
        evenOddHf: state.evenOddHf,
        multiSection: live.sections.length > 1,
      }
    },
    set(kind, variant, text) {
      if (deps.isLocked()) return 'the document is read-only; headers/footers cannot be edited'
      const live = hfState()
      const si = kind === 'header' ? 0 : Math.max(0, live.sections.length - 1)
      const target = hfCommitTarget(live, si, kind, variant)
      if (!target) return 'no document section is available'
      const next = applyHfText(resolveHf(live, si, kind, variant).value, text)
      if (variant === 'first' && !store.getSnapshot().titlePg) {
        store.setTitlePg(true)
        store.setTitlePgDirty(true)
      }
      if (variant === 'even' && !store.getSnapshot().evenOddHf) {
        store.setEvenOddHf(true)
        store.setEvenOddHfDirty(true)
      }
      store.setSectionHfEdits((edits) => ({ ...edits, [target.key]: next }))
      return null
    },
  }
  const extras: AiDocExtras = {
    styles: {
      list() {
        const parsed = deps.getContext().doc?.parsed
        const out = new Map<string, AiStyleInfo>()
        for (const s of parsed?.styles.values() ?? []) {
          if (s.linkedCharShell) continue
          out.set(s.styleId, {
            styleId: s.styleId,
            name: s.name,
            type: s.type,
            ...(s.basedOn ? { basedOn: s.basedOn } : {}),
            ...(s.headingLevel ? { headingLevel: s.headingLevel } : {}),
          })
        }
        const upserts = store.getSnapshot().styleUpserts
        for (const up of Object.values(upserts)) {
          const cur = out.get(up.styleId)
          const headingLevel = pendingHeadingLevel(
            up.styleId,
            (id) => upserts[id],
            (id) => parsed?.styles.get(id),
          )
          out.set(up.styleId, {
            styleId: up.styleId,
            name: up.name ?? cur?.name ?? up.styleId,
            type: cur?.type ?? up.type ?? 'paragraph',
            ...(up.basedOn === undefined
              ? cur?.basedOn
                ? { basedOn: cur.basedOn }
                : {}
              : up.basedOn
                ? { basedOn: up.basedOn }
                : {}),
            ...(headingLevel ? { headingLevel } : {}),
            pending: true,
          })
        }
        return [...out.values()]
      },
      upsert(up) {
        if (deps.isLocked()) return 'the document is read-only; styles cannot be edited'
        if (!deps.getContext().doc) return 'no document is open'
        store.setStyleUpserts((current) => {
          const prev = current[up.styleId]
          return {
            ...current,
            [up.styleId]: prev
              ? {
                  ...prev,
                  ...up,
                  pPr: up.pPr || prev.pPr ? { ...prev.pPr, ...up.pPr } : undefined,
                  rPr: up.rPr || prev.rPr ? { ...prev.rPr, ...up.rPr } : undefined,
                }
              : up,
          }
        })
        return null
      },
    },
    watermark: {
      current: () => store.getSnapshot().watermark,
      set(spec) {
        if (deps.isLocked()) return 'the document is read-only; the watermark cannot be edited'
        if (!deps.getContext().doc) return 'no document is open'
        if (spec && 'image' in spec) {
          store.setWatermark(null)
          store.setWatermarkStyle(null)
          store.setWatermarkPicture(spec)
        } else if (spec) {
          const { text, ...style } = spec
          store.setWatermark(text)
          store.setWatermarkStyle(Object.keys(style).length > 0 ? style : null)
          store.setWatermarkPicture(null)
        } else {
          store.setWatermark(null)
          store.setWatermarkStyle(null)
          store.setWatermarkPicture(null)
        }
        store.setWatermarkDirty(true)
        return null
      },
    },
  }
  const pageSetup = (deferSave = false): AiPageSetupAccess => ({
    list() {
      const { editor, doc } = deps.getContext()
      const { sections } = store.getSnapshot()
      if (!editor || !doc || sections.length === 0) return []
      const pmDoc = editor.state.doc
      const ranges = sections.map(() => ({ first: -1, last: -1 }))
      for (let i = 0; i < pmDoc.childCount; i++) {
        const r = ranges[sectionIndexAtBlock(pmDoc, sections, i)]!
        if (r.first < 0) r.first = i
        r.last = i
      }
      return sections.map((sec, i) =>
        describeSection(sec, i, Math.max(ranges[i]!.first, 0), Math.max(ranges[i]!.last, 0)),
      )
    },
    current: (index) => store.getSnapshot().sections[index],
    set(index, resolved) {
      if (deps.isLocked()) return 'the document is read-only; the page setup cannot be changed'
      const { editor } = deps.getContext()
      const state = store.getSnapshot()
      const sec = state.sections[index]
      if (!sec || !editor) return `section ${index} does not exist`
      const nextXml = applyResolvedPageSetup(
        effectiveLayoutSectPr(state, index) ?? sec.sectPrXml,
        resolved,
      )
      const next: SectionInfo = {
        ...sec,
        settings: resolved.settings,
        titlePg: resolved.titlePg ?? sec.titlePg,
        ...(resolved.pgNum
          ? { pageNumberStart: resolved.pgNum.start, pageNumberFmt: resolved.pgNum.fmt }
          : {}),
      }
      if (sec.pendingBreak) {
        if (!patchPendingSectPr(editor, state.sections, index, nextXml))
          return `section ${index} was removed from the document`
        next.sectPrXml = nextXml
      }
      store.setSections((prev) => prev.map((s, i) => (i === index ? next : s)))
      if (sec.pendingBreak) return null
      if (state.sections.length <= 1 || index === state.sections.length - 1) {
        store.setSection(resolved.settings)
        store.setSectionDirty(true)
        if (resolved.titlePg !== undefined) {
          store.setTitlePg(resolved.titlePg)
          store.setTitlePgDirty(true)
        }
        if (resolved.pgNum) store.setPgNumEdit(resolved.pgNum)
      } else {
        store.setSectionsDirty((d) => (d.includes(index) ? d : [...d, index]))
        if (resolved.pgNum)
          store.setPgNumDirtySections((d) => (d.includes(index) ? d : [...d, index]))
      }
      return null
    },
    insertBreak(type, afterBlockIndex) {
      if (deps.isLocked()) return 'the document is read-only; a section break cannot be inserted'
      const { editor, doc } = deps.getContext()
      if (!editor || !doc) return 'no document is open'
      const state = store.getSnapshot()
      const pmDoc = editor.state.doc
      const ownerIdx = sectionIndexAtBlock(pmDoc, state.sections, Math.max(afterBlockIndex, 0))
      const owner = state.sections[ownerIdx]
      const copyXml = effectiveLayoutSectPr(state, ownerIdx)
      if (!owner || !copyXml) return 'no document section is available'
      deps.insertBreak(type, afterBlockIndex, {
        sectPr: copyXml,
        sections: state.sections,
        deferSave,
      })
      let lastDocx: number | null = null
      for (let i = Math.min(afterBlockIndex, pmDoc.childCount - 1); i >= 0; i--) {
        const di = pmDoc.child(i).attrs?.docxIndex as number | null | undefined
        if (di !== null && di !== undefined) {
          lastDocx = di
          break
        }
      }
      const created: SectionInfo = {
        ...owner,
        pendingBreak: true,
        sectPrXml: copyXml,
        lastBlockIndex: lastDocx ?? owner.firstBlockIndex - 1,
      }
      const remainderXml = applySectionStartType(copyXml, type)
      const ownerNext: SectionInfo = {
        ...owner,
        firstBlockIndex: created.lastBlockIndex + 1,
        startType: type,
        ...(owner.pendingBreak ? { sectPrXml: remainderXml } : {}),
      }
      store.setSections((prev) => {
        const next = [...prev]
        next.splice(ownerIdx, 0, created)
        if (next[ownerIdx + 1])
          next[ownerIdx + 1] = {
            ...next[ownerIdx + 1]!,
            firstBlockIndex: ownerNext.firstBlockIndex,
            startType: type,
          }
        return next
      })
      const shift = (d: number[]) => d.map((i) => (i >= ownerIdx ? i + 1 : i))
      store.setSectionsDirty(shift)
      store.setPgNumDirtySections(shift)
      return null
    },
  })
  return { hf, extras, pageSetup }
}
