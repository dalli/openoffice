import {
  createUserSkillsSkill,
  skillFileAnswerOf,
  skillTextAnswerOf,
  userSkillEntriesOf,
  type SkillProposal,
  type UserSkillsBridgeLike,
  type UserSkillText,
  type WorkspaceTemplateMode,
  type WorkspaceTemplateSnapshot,
} from '@genoffice/agent-core'
import {
  buildSkillRequest,
  initialSkillCardDraft,
  pickerSkillsOf,
  readSkillsForRequest,
  skillCardSavedRecord,
  skillCardSaveResultOf,
  skillCardText,
  skillTemplateText,
  skillChipOf,
  skillChipRef,
  skillDisplayName,
  skillPickerText,
  previewEndedOf,
  withSkillChip,
  type AiComposerSkills,
  type AiSkillsBridge,
  type SkillCardDraft,
  type SkillCardBridge,
  type SkillCardSaveInput,
  type SkillChipView,
  type SkillPickerItem,
  type SkillTemplateBridge,
  type TemplateDocument,
} from '@genoffice/ui'
import type { Lang } from '@genoffice/i18n'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import type { WorkspaceComposer } from './conversations'
import type { WorkspaceRunHost, WorkspaceRunPreparation } from './run-host'
import type { WorkspaceDocumentIdentity } from '../../../shared/workspace-ai-api'
import type { WorkspaceCapture } from '../../../main/workspace-ai/execution-controller'
import type { WorkspacePanelEntry } from '../../../shared/workspace-panel-api'

export type WorkspaceSkillsBridge = AiSkillsBridge & UserSkillsBridgeLike
export interface WorkspaceSkillCard {
  proposal: SkillProposal
  draft: SkillCardDraft
  status: 'open' | 'saved' | 'unsaved' | 'preview'
  target?: WorkspaceDocumentIdentity
  continuesIn?: { tabId: string; previewId: string }
  beforeSeq: number
  afterSeq?: number
  saving?: boolean
}
interface TemplateNote {
  key: string
  skillId: string
  title: string
  text: string
  chips: WorkspaceComposer['skills']
  tabId?: string
  opening?: boolean
  several?: boolean
}

/** One panel's skills. Drafts, reads, proposals and template offers keep their original conversation. */
export class WorkspaceUserSkills {
  private raw: unknown = []
  private loaded = false
  private closed = false
  private sequence = 0
  private revision = 0
  private readonly listeners = new Set<() => void>()
  private readonly notices = new Map<string, string>()
  private readonly cards = new Map<string, WorkspaceSkillCard[]>()
  private readonly templates = new Map<string, TemplateNote[]>()
  private readonly restored = new Map<string, number>()
  private readonly handing = new Set<string>()
  private readonly handoffs = new Map<string, { capture: WorkspaceCapture; skills: Set<string> }>()
  private readonly off: () => void
  private readonly offPreview: () => void

  constructor(
    readonly bridge: WorkspaceSkillsBridge,
    private readonly host: WorkspaceRunHost,
    private readonly lang: () => Lang,
    readonly templateBridge: SkillTemplateBridge | null = null,
    private readonly mcpServerUsable: (conversationId: string, server: string) => boolean = () =>
      false,
    private readonly readTemplate?: (
      document: WorkspaceDocumentIdentity,
      mode: WorkspaceTemplateMode,
    ) => Promise<WorkspaceTemplateSnapshot | null>,
    private readonly openTemplateFor?: (
      id: string,
      input: Parameters<SkillTemplateBridge['openTemplate']>[0],
    ) => Promise<unknown>,
  ) {
    this.off = bridge.onChanged(() => void this.refresh().catch(() => {}))
    this.offPreview =
      templateBridge?.onPreviewEnded?.((raw) => {
        const ended = previewEndedOf(raw)
        if (!ended || this.closed) return
        for (const [id, cards] of this.cards)
          for (const card of cards) {
            if (card.status !== 'preview' || card.continuesIn?.previewId !== ended.previewId)
              continue
            if (ended.outcome === 'saved' && ended.skillId)
              void this.saved(id, card, ended.skillId, {
                ...card.draft,
                name: ended.title || card.draft.name,
              })
            else {
              card.status = 'unsaved'
              this.notify()
            }
          }
      }) ?? (() => {})
    void this.refresh().catch(() => {})
  }
  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  version = () => this.revision
  private notify() {
    if (this.closed) return
    this.revision++
    for (const listener of [...this.listeners]) listener()
  }
  async refresh(): Promise<void> {
    const sequence = ++this.sequence
    const raw = await this.bridge.list()
    if (this.closed || sequence !== this.sequence) return
    this.raw = raw
    this.loaded = true
    this.notify()
  }
  items(): SkillPickerItem[] {
    return pickerSkillsOf(this.raw, this.lang(), 'panel')
  }
  available = (id: string) => !this.loaded || this.items().some((item) => item.id === id)
  chips(id: string): SkillChipView[] {
    const items = new Map(this.items().map((item) => [item.id, item]))
    return this.host.conversations.state(id).composer.skills.map((chip) => {
      const item = items.get(chip.id)
      return item ? skillChipOf(item) : { ...chip, missing: this.loaded }
    })
  }
  composer(id: string): AiComposerSkills {
    return {
      lang: this.lang(),
      items: this.items(),
      chips: this.chips(id),
      notice: this.notices.get(id) ?? null,
      onOpen: () => void this.refresh().catch(() => {}),
      onManage: () => void this.bridge.openSettings().catch(() => {}),
      onPick: (item) => {
        const composer = this.host.conversations.state(id).composer
        this.host.conversations.edit(id, {
          composer: {
            ...composer,
            skills: withSkillChip(composer.skills, skillChipOf(item)).map(skillChipRef),
          },
        })
        this.notices.delete(id)
        this.notify()
      },
      onRemove: (skillId) => {
        const composer = this.host.conversations.state(id).composer
        this.host.conversations.edit(id, {
          composer: {
            ...composer,
            skills: composer.skills.filter((chip) => chip.id !== skillId),
          },
        })
        this.notices.delete(id)
        this.notify()
      },
    }
  }
  private refuse(id: string, key: 'missingNotice' | 'tooLarge'): never {
    const text = skillPickerText(this.lang(), key)
    this.notices.set(id, text)
    this.notify()
    throw new Error(text)
  }

  /** Main already opened this template's copy. Its chips must not open another copy. */
  acceptHandoff(
    id: string,
    capture: WorkspaceCapture,
    entry: Extract<WorkspacePanelEntry, { kind: 'template-handoff' }>,
  ): boolean {
    if (this.closed || capture.binding.document.kind !== 'docs') return false
    this.handoffs.set(id, {
      capture: structuredClone(capture),
      skills: new Set(entry.chips.map((chip) => chip.id)),
    })
    return true
  }
  private handedSkill(id: string, input: WorkspaceComposer, skillId: string): boolean {
    const handoff = this.handoffs.get(id)
    const target = input.target
    const bound = handoff?.capture.binding
    return (
      !!bound &&
      !!target &&
      handoff!.skills.has(skillId) &&
      bound.revision === target.revision &&
      (['windowId', 'tabId', 'documentId', 'generation', 'kind'] as const).every(
        (key) => bound.document[key] === target.document[key],
      )
    )
  }

  /** Template copies use the existing explicit handoff; their editor runs its own send gate. */
  async send(id: string, send: () => Promise<unknown>): Promise<void> {
    if (this.closed || this.handing.has(id)) return
    const input = this.host.conversations.state(id).composer
    if (!input.skills.length) {
      await send()
      return
    }
    this.handing.add(id)
    try {
      await this.refresh()
      if (this.closed) return
      const items = new Map(this.items().map((item) => [item.id, item]))
      if (input.skills.some((chip) => !items.has(chip.id))) this.refuse(id, 'missingNotice')
      {
        const templateChips = input.skills.filter((chip) => {
          const item = items.get(chip.id)!
          return item.template || item.templateMissing
        })
        if (templateChips.some((chip) => this.handoffs.get(id)?.skills.has(chip.id))) {
          if (templateChips.some((chip) => !this.handedSkill(id, input, chip.id)))
            this.refuse(id, 'missingNotice')
          if (JSON.stringify(this.host.conversations.state(id).composer) === JSON.stringify(input))
            await send()
          return
        }
      }
      const template = input.skills
        .map((chip) => items.get(chip.id)!)
        .find((s) => s.template || s.templateMissing)
      if (!template) {
        if (JSON.stringify(this.host.conversations.state(id).composer) === JSON.stringify(input))
          await send()
        return
      }
      if (!this.templateBridge) this.refuse(id, 'missingNotice')
      await this.openTemplate(
        id,
        {
          key: crypto.randomUUID(),
          skillId: template.id,
          title: template.title,
          text: input.text,
          chips: input.skills,
          several:
            input.skills.filter(
              (chip) => items.get(chip.id)?.template || items.get(chip.id)?.templateMissing,
            ).length > 1,
        },
        true,
        input,
      )
    } catch {
      // Only translated skill notices reach the UI, never bridge error text or paths.
      if (!this.closed && !this.notices.has(id)) this.refusalNotice(id)
    } finally {
      this.handing.delete(id)
    }
  }
  private refusalNotice(id: string) {
    this.notices.set(id, skillTemplateText(this.lang(), 'openFailed'))
    this.notify()
  }
  templateNotes(id: string): readonly TemplateNote[] {
    return this.templates.get(id) ?? []
  }
  async openTemplate(id: string, note: TemplateNote, chosen = false, input?: WorkspaceComposer) {
    if (this.closed || !this.templateBridge || note.opening || note.tabId) return
    note.opening = true
    this.notify()
    try {
      const request = {
        skillId: note.skillId,
        text: note.text,
        chips: note.chips,
      }
      const raw = await (this.openTemplateFor
        ? this.openTemplateFor(id, request)
        : this.templateBridge.openTemplate(request))
      if (this.closed) return
      const result = raw as { ok?: boolean; tabId?: string }
      if (!result?.ok || typeof result.tabId !== 'string') {
        this.notices.set(
          id,
          skillTemplateText(this.lang(), 'templateMissing', { title: note.title }),
        )
        return
      }
      note.tabId = result.tabId
      if (chosen) {
        const notes = this.templates.get(id) ?? []
        notes.push(note)
        this.templates.set(id, notes)
        const now = this.host.conversations.state(id).composer
        // A late native handoff must not clear a newer draft in either conversation.
        if (input && JSON.stringify(now) === JSON.stringify(input))
          this.host.conversations.edit(id, { composer: { ...now, text: '', skills: [] } })
      }
      this.notices.delete(id)
    } catch {
      if (!this.closed) this.refusalNotice(id)
    } finally {
      note.opening = false
      this.notify()
    }
  }
  proposals(id: string): readonly WorkspaceSkillCard[] {
    return this.cards.get(id) ?? []
  }
  /** Only a proposal's original target supplies a template; model requests and histories never receive it. */
  document(card: WorkspaceSkillCard): TemplateDocument | null {
    const target = card.target
    if (!target || target.kind !== 'docs' || !this.readTemplate) return null
    let info: WorkspaceTemplateSnapshot | null = null
    let copy: WorkspaceTemplateSnapshot | null = null
    const read = async (mode: WorkspaceTemplateMode) => {
      if (this.closed || card.status !== 'open' || this.host.busy())
        throw new Error('Template unavailable')
      const snapshot = await this.readTemplate!(target, mode)
      if (this.closed || card.status !== 'open' || this.host.busy() || !snapshot)
        throw new Error('Template unavailable')
      info = { ...snapshot, bytes: null }
      if (mode !== 'info') copy = info
      return snapshot
    }
    return {
      get fileName() {
        return copy?.fileName ?? info?.fileName ?? ''
      },
      get title() {
        return copy?.title ?? info?.title ?? ''
      },
      redacted: async () => (await read('info')).redacted || copy?.redacted === true,
      bytes: async (options) =>
        (await read(options?.allowRedacted === false ? 'as-is' : 'preview')).bytes,
    }
  }
  previewOpened(card: WorkspaceSkillCard, opened: { tabId: string; previewId: string }) {
    if (this.closed || card.status !== 'open') return
    card.status = 'preview'
    card.continuesIn = opened
    this.notify()
  }
  draft(card: WorkspaceSkillCard, draft: SkillCardDraft) {
    card.draft = draft
  }
  cancel(card: WorkspaceSkillCard) {
    if (card.saving) return
    card.status = 'unsaved'
    this.notify()
  }
  /** The conversation owns an in-flight save, including while its card view is unmounted. */
  async save(
    id: string,
    card: WorkspaceSkillCard,
    bridge: SkillCardBridge,
    input: SkillCardSaveInput,
  ): Promise<unknown> {
    if (this.closed || card.saving || card.status !== 'open') return { ok: false, reason: 'failed' }
    const draft = structuredClone(card.draft)
    card.saving = true
    this.notify()
    try {
      const result = await bridge.saveFromCard(input)
      const answer = skillCardSaveResultOf(result)
      if (answer.ok) await this.saved(id, card, answer.id, draft)
      return result
    } finally {
      card.saving = false
      this.notify()
    }
  }
  async saved(id: string, card: WorkspaceSkillCard, skillId: string, draft: SkillCardDraft) {
    if (this.closed) return
    card.draft = draft
    card.status = 'saved'
    this.notify()
    await this.host.appendNote(id, {
      role: 'assistant',
      text: '',
      tools: [skillCardSavedRecord(card.proposal.name || card.proposal.title, skillId)],
    })
  }

  async prepare(
    input: WorkspaceComposer,
    run: WorkspaceStartedRun,
  ): Promise<WorkspaceRunPreparation> {
    const id = run.binding.conversationId
    let disposed = false
    let started = false
    const runCards: WorkspaceSkillCard[] = []
    const check = () => {
      if (this.closed || disposed || run.signal.aborted) throw new Error('Skill execution stopped')
    }
    const read = async <T>(action: () => Promise<T>): Promise<T> => {
      check()
      const answer = await action()
      check()
      return answer
    }
    try {
      await read(() => this.refresh())
    } catch {
      check()
      this.refuse(id, 'missingNotice')
    }
    const items = new Map(this.items().map((item) => [item.id, item]))
    if (
      input.skills.some((chip) => {
        const item = items.get(chip.id)
        return (
          !item ||
          ((item.template || item.templateMissing) && !this.handedSkill(id, input, chip.id))
        )
      })
    )
      this.refuse(id, 'missingNotice')
    const revision = this.host.conversations.historyRevision(id)
    const history =
      this.restored.get(id) === revision
        ? []
        : this.host.conversations.state(id).messages.flatMap((m) => m.skills ?? [])
    const ids = [...new Set([...input.skills, ...history].map((chip) => chip.id))]
    const selected = await read(() =>
      readSkillsForRequest({ aiRead: (id) => read(() => this.bridge.aiRead(id)) }, ids),
    )
    if (input.skills.some((chip) => selected.missing.includes(chip.id)))
      this.refuse(id, 'missingNotice')
    const built = buildSkillRequest(selected.skills, undefined, {
      mcpServerUsable: (server) => this.mcpServerUsable(id, server),
    })
    if (!built.ok) this.refuse(id, 'tooLarge')
    this.notices.delete(id)
    this.notify()
    let entries = userSkillEntriesOf(this.raw)
    const used = (skill: UserSkillText) => {
      check()
      if (!skill.template) return
      const notes = this.templates.get(id) ?? []
      notes.push({
        key: crypto.randomUUID(),
        skillId: skill.id,
        title: skill.title,
        text: input.text,
        chips: [{ id: skill.id, source: skill.id.split(':')[0]!, title: skill.title }],
      })
      this.templates.set(id, notes)
      this.notify()
    }
    const capability = createUserSkillsSkill({
      list: () => entries,
      reload: async () => {
        await read(() => this.refresh())
        entries = userSkillEntriesOf(this.raw)
      },
      read: async (id) => skillTextAnswerOf(await read(() => this.bridge.aiRead(id))),
      readFile: async (id, path) =>
        skillFileAnswerOf(await read(() => this.bridge.aiReadFile(id, path))),
      messages: () => [{ role: 'user', text: input.text + built.text }],
      onUsed: used,
      onProposal: (proposal) => {
        check()
        const cards = this.cards.get(id) ?? []
        const unique = { ...proposal, key: `${run.binding.runId}:${proposal.key}` }
        const card: WorkspaceSkillCard = {
          proposal: unique,
          draft: initialSkillCardDraft(unique),
          status: 'open',
          target: structuredClone(run.binding.target.document),
          beforeSeq: this.host.conversations.state(id).messages.at(-1)?.seq ?? -1,
        }
        cards.push(card)
        runCards.push(card)
        this.cards.set(id, cards)
        this.notify()
      },
      labels: {
        used: (title) => skillCardText(this.lang(), 'skillUsed', { title }),
        proposed: (title) => skillCardText(this.lang(), 'skillProposed', { title }),
      },
      displayTitle: (skill) =>
        skillDisplayName({
          name: skill.name,
          title: skill.source === 'builtin' ? '' : skill.title,
        }),
      mcpServerUsable: (server) => this.mcpServerUsable(id, server),
    })
    return {
      project: run.binding.project,
      userText: input.text + built.text,
      skills: [
        {
          id: capability.id,
          get systemPrompt() {
            return capability.systemPrompt
          },
          tools: capability.tools,
          executeTool: async (call, signal) => {
            check()
            if (signal?.aborted) throw new Error('Skill execution stopped')
            const answer = await capability.executeTool(call, signal)
            check()
            if (signal?.aborted) throw new Error('Skill execution stopped')
            return answer
          },
        },
      ],
      onStart: () => {
        started = true
      },
      dispose: () => {
        if (started && !this.closed) {
          this.restored.set(id, this.host.conversations.historyRevision(id))
          for (const card of runCards)
            card.afterSeq = this.host.conversations.state(id).messages.at(-1)?.seq
          this.notify()
        }
        disposed = true
      },
    }
  }
  dispose() {
    this.closed = true
    this.off()
    this.offPreview()
    this.listeners.clear()
    this.cards.clear()
    this.templates.clear()
    this.notices.clear()
    this.restored.clear()
    this.handoffs.clear()
  }
}
