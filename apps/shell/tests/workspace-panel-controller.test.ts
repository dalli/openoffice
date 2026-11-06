import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  WorkspacePanelController,
  type WorkspacePanelDeps,
} from '../src/renderer/src/workspace-ai/panel-controller'
import { makeWorkspaceHost, until } from './helpers/workspace-host-fixture'

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach((off) => off()))
const source = (id: string) => ({
  origin: { kind: 'workspace' as const, conversationId: id },
  title: id,
})
function fixture(patch: Partial<WorkspacePanelDeps> = {}) {
  const f = makeWorkspaceHost()
  f.addEditor('A')
  const b = f.addEditor('B')
  let next = 0
  const capture = vi.fn((id: string) => f.execution.capture(id))
  const reference = vi.fn(async (id: string, scope: 'document' | 'selection') => {
    const { id: refId, binding } = await f.execution.readReference(id, { scope })
    return { id: refId, binding }
  })
  const controller = new WorkspacePanelController({
    host: f.host,
    capture,
    refresh: (id) => f.execution.refresh(id),
    reference,
    create: async () => source(`new-${++next}`),
    ...patch,
  })
  controller.catalog(
    [
      { id: 'A', title: 'Duplicate', kind: 'docs', active: true, available: true },
      { id: 'B', title: 'Duplicate', kind: 'docs', available: true },
      { id: 'unloaded', title: 'Unloaded recovery tab', kind: 'docs', available: false },
    ],
    [source('P'), source('Q')],
  )
  cleanups.push(() => {
    controller.dispose()
    f.dispose()
  })
  return { ...f, controller, capture, reference, b }
}

describe('workspace panel selection ownership', () => {
  it('installs handed template chips in a new conversation, admits once, and preserves other drafts', async () => {
    const installed = vi.fn(() => true)
    const f = fixture({ handoff: installed })
    await f.controller.create()
    const previous = f.host.conversations.selectedId!
    f.host.conversations.edit(previous, {
      composer: { ...f.host.conversations.state(previous).composer, text: 'KEEP' },
    })
    const capture = await f.execution.capture('B')
    const entry = {
      kind: 'template-handoff' as const,
      id: 'handoff',
      text: 'Fill this copy',
      chips: [{ id: 'app:one', source: 'app', title: 'One' }],
    }
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    const accept = vi.fn(async (supported = true) => {
      expect(supported).toBe(true)
      const id = f.host.conversations.selectedId!
      expect(f.host.conversations.state(id).composer.skills).toEqual(entry.chips)
      expect(f.controller.canSend(id)).toBe(false)
      return true
    })
    await f.controller.enter(capture, entry, accept)
    const id = f.host.conversations.selectedId!
    expect(installed).toHaveBeenCalledOnce()
    expect(accept).toHaveBeenCalledOnce()
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(previous).composer.text).toBe('KEEP')
    expect(f.host.conversations.state(id).composer.text).toBe(entry.text)
    f.setGate({ kind: 'run', choice: { connected: true, provider: 'claude-code', metered: false } })
    await f.controller.send(id)
    await until(() => f.requests.length === 1)
    expect(f.host.conversations.state(id).messages[0]!.skills).toEqual(entry.chips)
    f.callbacks.get(id)!.at(-1)!.onDone()
  })

  it('does not send rejected, interrupted, or changed template transfers', async () => {
    const f = fixture({ handoff: () => true })
    const entry = { kind: 'template-handoff' as const, id: 'handoff', text: 'Fill', chips: [] }
    await f.controller.enter(await f.execution.capture('B'), entry, async () => false)
    expect(f.controller.canSend(f.host.conversations.selectedId!)).toBe(false)
    expect(f.host.conversations.state(f.host.conversations.selectedId!).composer.text).toBe('')
    await f.controller.enter(await f.execution.capture('B'), entry, async () => {
      throw new Error('Lost acknowledgement')
    })
    expect(f.controller.canSend(f.host.conversations.selectedId!)).toBe(false)
    expect(f.host.conversations.state(f.host.conversations.selectedId!).composer.text).toBe('Fill')
    await f.controller.enter(await f.execution.capture('B'), entry, async () => {
      const id = f.host.conversations.selectedId!
      f.host.conversations.edit(id, {
        composer: { ...f.host.conversations.state(id).composer, text: 'NEW DRAFT' },
      })
      return true
    })
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(f.host.conversations.selectedId!).composer.text).toBe(
      'NEW DRAFT',
    )
  })

  it('keeps an interrupted acknowledgement from leaving a sendable batch behind', async () => {
    const f = fixture()
    const capture = await f.execution.captureEditBatch('B', { documentId: 'B', revision: '1' })
    await f.controller.enter(
      capture,
      {
        kind: 'edit-batch',
        text: 'BATCH',
        displayText: 'Queued edit',
        documentId: 'B',
        revision: '1',
      },
      async () => {
        throw new Error('Disconnected')
      },
    )
    const id = f.host.conversations.selectedId!
    expect(f.controller.selection(id).submitting).toBe(false)
    expect(f.controller.canSend(id)).toBe(false)
    expect(f.requests).toEqual([])
  })
  it('transfers a batch only after its draft is installed, uses the send gate, and stores its summary', async () => {
    const f = fixture()
    await f.controller.create()
    const previous = f.host.conversations.selectedId!
    f.host.conversations.edit(previous, {
      composer: { ...f.host.conversations.state(previous).composer, text: 'KEEP_DRAFT' },
    })
    const capture = await f.execution.captureEditBatch('B', { documentId: 'B', revision: '1' })
    const entry = {
      kind: 'edit-batch' as const,
      text: 'MODEL_BATCH_INSTRUCTIONS',
      displayText: 'Two queued edits',
      documentId: 'B',
      revision: '1',
    }
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    const accept = vi.fn(async () => {
      const id = f.host.conversations.selectedId!
      expect(f.host.conversations.state(id).composer.text).toBe(entry.displayText)
      expect(f.controller.canSend(id)).toBe(false)
      return true
    })
    await f.controller.enter(capture, entry, accept)
    const id = f.host.conversations.selectedId!
    expect(accept).toHaveBeenCalledOnce()
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(previous).composer.text).toBe('KEEP_DRAFT')
    expect(f.host.conversations.state(id).composer.text).toBe(entry.displayText)
    f.setGate({ kind: 'run', choice: { connected: true, provider: 'claude-code', metered: false } })
    await f.controller.send(id)
    await until(() => f.requests.length === 1)
    expect(JSON.stringify(f.requests[0])).toContain(entry.text)
    expect(f.host.conversations.state(id).messages[0]!.text).toBe(entry.displayText)
    f.callbacks.get(id)!.at(-1)!.onDelta('Finished batch')
    f.callbacks.get(id)!.at(-1)!.onDone()
    await until(() => f.append.mock.calls.length === 2)
    expect(JSON.stringify(f.append.mock.calls)).toContain(entry.displayText)
    expect(JSON.stringify(f.append.mock.calls)).not.toContain(entry.text)
  })

  it('never refreshes a held batch onto a changed document and never runs an unacknowledged batch', async () => {
    const f = fixture()
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    const capture = await f.execution.captureEditBatch('B', { documentId: 'B', revision: '1' })
    const entry = {
      kind: 'edit-batch' as const,
      text: 'BATCH',
      displayText: 'Queued edit',
      documentId: 'B',
      revision: '1',
    }
    await f.controller.enter(capture, entry, async () => false)
    expect(f.host.conversations.state(f.host.conversations.selectedId!).composer.text).toBe('')
    await f.controller.enter(capture, entry, async () => true)
    const id = f.host.conversations.selectedId!
    f.b.manual()
    f.setGate({ kind: 'run', choice: { connected: true, provider: 'claude-code', metered: false } })
    await f.controller.send(id)
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(id).error).toBeTruthy()
    expect(f.host.conversations.state(id).composer.text).toBe(entry.displayText)
  })
  it('pins an editor preset to a new conversation and leaves the previous draft and references intact', async () => {
    const f = fixture()
    await f.controller.create()
    const before = f.host.conversations.selectedId!
    await f.controller.addReference(before, 'B', 'document')
    const composer = { ...f.host.conversations.state(before).composer, text: 'KEEP_DRAFT' }
    f.host.conversations.edit(before, { composer })
    const capture = await f.execution.capture('B')
    f.capture.mockClear()
    await f.controller.enter(capture, { kind: 'preset', text: 'EDITOR_PRESET', autoRun: false })
    const id = f.host.conversations.selectedId!
    expect(id).not.toBe(before)
    expect(f.controller.selection(id).target).toEqual(capture)
    expect(f.host.conversations.state(id).composer.text).toBe('EDITOR_PRESET')
    expect(f.host.conversations.state(before).composer).toEqual(composer)
    expect(f.capture).not.toHaveBeenCalled()
    expect(f.requests).toEqual([])
    await f.controller.enter(await f.execution.capture('A'), { kind: 'open' })
    expect(f.host.conversations.selectedId).toBe(id)
    expect(f.controller.selection(id).target).toEqual(capture)
  })

  it('uses the existing send gate for an immediate editor preset and keeps a blocked request as a draft', async () => {
    const f = fixture({})
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    const capture = await f.execution.capture('B')
    await f.controller.enter(capture, { kind: 'preset', text: 'BLOCKED_PRESET', autoRun: true })
    const id = f.host.conversations.selectedId!
    expect(f.host.conversations.state(id).composer.text).toBe('BLOCKED_PRESET')
    expect(f.requests).toEqual([])
    expect(f.controller.selection(id).target?.binding.document.tabId).toBe('B')
  })
  it('fills only the chosen conversation draft and preserves its pinned target, scope, attachments and references', async () => {
    const f = fixture()
    await f.controller.create()
    const p = f.host.conversations.selectedId!
    await f.controller.addReference(p, 'B', 'document')
    const before = f.host.conversations.state(p).composer
    before.attachments = [{ name: 'fixture.md', path: '/fixture.md', size: 10 }]
    f.host.conversations.edit(p, { composer: before })
    const target = f.controller.selection(p).target!
    await f.controller.create()
    const q = f.host.conversations.selectedId!
    f.capture.mockClear()
    expect(f.controller.preset(p, 'polish this selection', target.binding.document)).toBe(true)
    expect(f.host.conversations.state(p).composer).toEqual({
      ...before,
      text: 'polish this selection',
    })
    expect(f.host.conversations.state(q).composer.text).toBe('')
    expect(f.host.conversations.selectedId).toBe(q)
    expect(f.controller.selection(p).target).toEqual(target)
    expect(f.capture).not.toHaveBeenCalled()
    expect(f.requests).toEqual([])
    await f.controller.target(p, 'B')
    expect(f.controller.preset(p, 'late map question', target.binding.document)).toBe(false)
    expect(f.host.conversations.state(p).composer.text).toBe('polish this selection')
    const current = f.controller.selection(p).target!.binding.document
    expect(
      f.controller.preset(p, 'wrong generation', {
        ...current,
        generation: current.generation + 1,
      }),
    ).toBe(false)
    f.controller.catalog([], [])
    expect(f.controller.preset(p, 'closed')).toBe(false)
  })
  it('renews only the target between consecutive runs and after undo in the same conversation', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.addReference(id, 'B', 'document')
    const reference = f.controller.selection(id).references[0]!
    const edit = (text: string) =>
      f.host.conversations.edit(id, {
        composer: { ...f.host.conversations.state(id).composer, text },
      })
    for (let index = 0; index < 3; index++) {
      edit(`request ${index}`)
      await f.controller.send(id)
      await until(() => f.requests.length === index * 2 + 1)
      const cb = f.callbacks.get(id)!.at(-1)!
      cb.onToolCall({ id: `write-${index}`, name: 'write', input: { text: `changed ${index}` } })
      cb.onDone()
      await until(() => f.requests.length === index * 2 + 2)
      f.callbacks.get(id)!.at(-1)!.onDelta(`done ${index}`)
      f.callbacks.get(id)!.at(-1)!.onDone()
      await until(() => !f.host.busy())
      if (index === 1) await f.host.rollback(id, f.host.rollbackCards(id).at(-1)!.id)
      expect(f.controller.selection(id).references[0]).toEqual(reference)
    }
    expect(f.host.conversations.state(id).error).toBeNull()
    expect(f.reference).toHaveBeenCalledOnce()
    expect(f.capture).toHaveBeenCalledTimes(2)
  })

  it('keeps a changed reference frozen and refuses sending until it is explicitly selected again', async () => {
    const f = fixture()
    const referenceEditor = f.addEditor('B')
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.addReference(id, 'B', 'document')
    const before = f.controller.selection(id).references
    referenceEditor.manual()
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'use reference' },
    })
    await f.controller.send(id)
    expect(f.requests).toEqual([])
    expect(f.controller.selection(id).references).toEqual(before)
    expect(f.host.conversations.state(id).composer.text).toBe('use reference')
    await f.controller.addReference(id, 'B', 'document')
    await f.controller.send(id)
    await until(() => f.requests.length === 1)
  })

  it('discards a late renewal after retargeting or disposal and retains the unsent draft', async () => {
    let resolve!: (value: Awaited<ReturnType<WorkspacePanelDeps['refresh']>>) => void
    const f = fixture({
      refresh: () =>
        new Promise((done) => {
          resolve = done
        }),
    })
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    const original = f.controller.selection(id).target!
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'keep draft' },
    })
    const sending = f.controller.send(id)
    await f.controller.target(id, 'B')
    resolve(original)
    await sending
    expect(f.host.conversations.state(id).composer.target?.document.tabId).toBe('B')
    expect(f.host.conversations.state(id).composer.text).toBe('keep draft')
    const disposed = f.controller.send(id)
    f.controller.dispose()
    resolve(original)
    await disposed
    expect(f.requests).toEqual([])
  })

  it('keeps a newly created source unique when the main catalog refresh arrives before creation finishes', async () => {
    const f = fixture({
      create: async () => {
        f.controller.catalog(f.controller.documentList(), [
          ...f.controller.sourceList(),
          source('new'),
        ])
        return source('new')
      },
    })
    await f.controller.create()
    expect(
      f.controller
        .sourceList()
        .filter((item) => item.origin.kind === 'workspace' && item.origin.conversationId === 'new'),
    ).toHaveLength(1)
  })

  it('opens source histories without a model call or a document read; only new conversations propose the focused target', async () => {
    const f = fixture()
    await f.controller.open(source('P'))
    const p = f.host.conversations.selectedId!
    expect(f.controller.selection(p).target).toBeNull()
    expect(f.capture).not.toHaveBeenCalled()
    await f.controller.create()
    const q = f.host.conversations.selectedId!
    expect(f.controller.selection(q).target?.binding.document.tabId).toBe('A')
    f.controller.select(p)
    expect(f.host.conversations.selectedId).toBe(p)
    expect(f.controller.selection(p).target).toBeNull()
    expect(f.reference).not.toHaveBeenCalled()
    expect(f.requests).toEqual([])
  })

  it('keeps draft/selection/reference/scroll state per conversation and never follows document focus', async () => {
    const f = fixture()
    await f.controller.create()
    const p = f.host.conversations.selectedId!
    await f.controller.addReference(p, 'B', 'document')
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'draft P' },
      scrollTop: 130,
    })
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    await f.controller.target(q, 'B')
    f.controller.catalog(
      f.controller.documentList().map((doc) => ({ ...doc, active: doc.id === 'B' })),
      f.controller.sourceList(),
    )
    f.controller.select(p)
    expect(f.host.conversations.state(p).composer.text).toBe('draft P')
    expect(f.host.conversations.state(p).scrollTop).toBe(130)
    expect(f.controller.selection(p).target?.binding.document.tabId).toBe('A')
    expect(f.controller.selection(p).references).toMatchObject([
      { scope: 'document', binding: { document: { tabId: 'B' } } },
    ])
    expect(f.host.conversations.state(q).composer.references).toEqual([])
    expect(f.reference).toHaveBeenCalledTimes(1)
    expect(f.requests).toEqual([])
  })

  it('rejects unavailable recovery tabs without capturing or loading them and retains the draft when a target closes', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.target(id, 'unloaded')
    expect(f.controller.selection(id).target).toBeNull()
    expect(f.host.conversations.state(id).composer.target).toBeNull()
    expect(f.capture).toHaveBeenCalledTimes(1)
    expect(f.controller.selection(id).error).toBe(true)
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'keep this' },
    })
    f.controller.catalog(
      f.controller.documentList().filter((doc) => doc.id !== 'A'),
      f.controller.sourceList(),
    )
    expect(f.controller.canSend(id)).toBe(false)
    await f.controller.send(id)
    expect(f.host.conversations.state(id).composer.text).toBe('keep this')
    expect(f.requests).toEqual([])
  })

  it('does not begin another read or history operation after the window owner disposes the panel controller', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    f.capture.mockClear()
    f.load.mockClear()
    f.controller.dispose()
    await f.controller.target(id, 'B')
    await f.controller.addReference(id, 'B', 'document')
    await f.controller.open(source('Q'))
    await f.controller.create()
    expect(f.capture).not.toHaveBeenCalled()
    expect(f.reference).not.toHaveBeenCalled()
    expect(f.load).not.toHaveBeenCalled()
    expect(f.controller.canSend(id)).toBe(false)
  })

  it('does not let a slow history open replace a more recent explicit selection', async () => {
    const f = fixture()
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    let resolve!: (value: { messages: []; foreign: boolean }) => void
    f.load.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const pending = f.controller.open(source('P'))
    await until(() => !!resolve)
    f.controller.select(q)
    resolve({ messages: [], foreign: false })
    await pending
    expect(f.host.conversations.selectedId).toBe(q)
  })

  it('ignores an older target capture and routes a late reference to its original conversation', async () => {
    let resolve!: (value: Awaited<ReturnType<WorkspacePanelDeps['capture']>>) => void
    const f = fixture()
    await f.controller.open(source('P'))
    const p = f.host.conversations.selectedId!
    const a = await f.execution.capture('A')
    f.capture.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const old = f.controller.target(p, 'A')
    await f.controller.target(p, 'B')
    resolve(a)
    await old
    expect(f.controller.selection(p).target?.binding.document.tabId).toBe('B')
    let referenceDone!: () => void
    const wait = new Promise<void>((done) => {
      referenceDone = done
    })
    f.reference.mockImplementationOnce(async (id, scope) => {
      await wait
      const { id: refId, binding } = await f.execution.readReference(id, { scope })
      return { id: refId, binding }
    })
    const pending = f.controller.addReference(p, 'A', 'document')
    await until(() => f.reference.mock.calls.length > 0)
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    referenceDone()
    await pending
    expect(f.controller.selection(p).references).toHaveLength(1)
    expect(f.controller.selection(q).references).toEqual([])
    expect(f.controller.selection(p).pending).toBe(0)
    expect(f.requests).toEqual([])
  })

  it('uses only the requested reference scope and keeps failed reads out of the composer', async () => {
    const reference = vi.fn(async () => {
      throw new Error('Selection is not available')
    })
    const f = fixture({ reference })
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.addReference(id, 'B', 'selection')
    expect(reference).toHaveBeenCalledWith(expect.any(String), 'selection')
    expect(f.controller.selection(id).references).toEqual([])
    expect(f.host.conversations.state(id).composer.references).toEqual([])
    expect(f.controller.selection(id).error).toBe(true)
  })

  it('sends through the existing guarded host, preserves later drafts and prevents double submission during admission', async () => {
    const f = fixture()
    await f.controller.create()
    const p = f.host.conversations.selectedId!
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'edit A' },
    })
    const sending = f.controller.send(p)
    const duplicate = f.controller.send(p)
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'later draft' },
    })
    await Promise.all([sending, duplicate])
    await until(() => f.requests.length === 1)
    expect(f.host.conversations.state(p).composer.text).toBe('later draft')
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    f.callbacks.get(p)![0]!.onDelta('answer for P')
    f.callbacks.get(p)![0]!.onDone()
    await until(() => !f.host.busy())
    expect(f.host.conversations.state(p).messages.at(-1)?.text).toBe('answer for P')
    expect(f.host.conversations.state(q).messages).toEqual([])
  })

  it('keeps a blocked draft and never sends automatically after the gate becomes ready', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'keep this draft' },
    })
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    await f.controller.send(id)
    expect(f.host.conversations.state(id).composer.text).toBe('keep this draft')
    expect(f.requests).toEqual([])
  })
})

describe('queue entry navigation', () => {
  it('opens the marked document without changing an existing conversation draft or sending a model request', async () => {
    const f = fixture()
    await f.controller.open(source('P'))
    const previous = f.host.conversations.selectedId!
    f.host.conversations.edit(previous, {
      composer: { ...f.host.conversations.state(previous).composer, text: 'KEEP_DRAFT' },
    })
    await f.controller.enter(await f.execution.capture('B'), { kind: 'queue' })
    const queue = f.host.conversations.selectedId!
    expect(queue).not.toBe(previous)
    expect(f.controller.selection(queue).target?.binding.document.tabId).toBe('B')
    f.host.conversations.edit(queue, {
      composer: { ...f.host.conversations.state(queue).composer, text: 'KEEP_QUEUE_DRAFT' },
    })
    await f.controller.enter(await f.execution.capture('B'), { kind: 'queue' })
    expect(f.host.conversations.selectedId).toBe(queue)
    expect(f.host.conversations.state(queue).composer.text).toBe('KEEP_QUEUE_DRAFT')
    expect(f.host.conversations.state(previous).composer.text).toBe('KEEP_DRAFT')
    expect(f.requests).toEqual([])
  })
})

describe('project template handoff conversations', () => {
  it('creates a distinct project history and keeps the previous draft untouched before admission', async () => {
    const create = vi.fn(async (projectId?: string) =>
      projectId
        ? {
            origin: { kind: 'project' as const, projectId, chatId: 'new-chat' },
            title: 'Project copy',
          }
        : source('previous'),
    )
    const f = fixture({ create, handoff: () => true })
    await f.controller.create()
    const previous = f.host.conversations.selectedId!
    f.host.conversations.edit(previous, {
      composer: { ...f.host.conversations.state(previous).composer, text: 'KEEP' },
    })
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    await f.controller.enter(
      await f.execution.capture('B'),
      {
        kind: 'template-handoff',
        id: 'project-copy',
        text: 'Fill copy',
        chips: [],
        projectId: 'project-one',
      },
      async () => true,
    )
    expect(create).toHaveBeenLastCalledWith('project-one')
    const next = f.host.conversations.state(f.host.conversations.selectedId!)
    expect(next.origin).toEqual({ kind: 'project', projectId: 'project-one', chatId: 'new-chat' })
    expect(next.composer.text).toBe('Fill copy')
    expect(f.host.conversations.state(previous).composer.text).toBe('KEEP')
    expect(f.requests).toEqual([])
  })
})
