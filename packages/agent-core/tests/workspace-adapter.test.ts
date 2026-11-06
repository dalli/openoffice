import { describe, expect, it, vi } from 'vitest'
import {
  createWorkspaceDocumentAdapter,
  WORKSPACE_READ_MAX_CHARS,
  type WorkspaceAdapterState,
  type WorkspaceDocumentAdapterDeps,
} from '../src/workspace-adapter'
import type { AgentSkill } from '../src/skill'
import type { AgentToolCall, AgentToolDef } from '../src/types'

const definition = (name: string): AgentToolDef => ({
  name,
  description: name,
  inputSchema: {
    type: 'object',
    properties: { value: { type: 'string' } },
    required: ['value'],
  },
})
const call = (name = 'write', id = 'call'): AgentToolCall => ({
  id,
  name,
  input: { value: 'replacement' },
})
const deferred = <T>() => {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => (resolve = done))
  return { promise, resolve }
}

function fixture(overrides: Partial<WorkspaceDocumentAdapterDeps> = {}) {
  let state: WorkspaceAdapterState = {
    instanceId: 'document-A',
    revision: '1',
    selectionId: 'captured-range',
    readOnly: false,
  }
  const executeTool = vi.fn<AgentSkill['executeTool']>((request) => {
    if (request.name === 'write') state = { ...state, revision: String(+state.revision + 1) }
    return { output: 'done', summary: request.name, mutated: request.name === 'write' }
  })
  const skill: AgentSkill = {
    id: 'fixture',
    systemPrompt: 'local document',
    tools: [definition('read'), definition('write'), definition('network')],
    buildContext: () => 'current unsaved body',
    executeTool,
    verifyResponse: vi.fn(() => 'correction'),
  }
  const read = vi.fn(() => ({ text: 'current unsaved body', truncated: false }))
  const beforeFirstWrite = vi.fn()
  const adapter = createWorkspaceDocumentAdapter({
    getState: () => state,
    read,
    createSkill: () => skill,
    toolAccess: { read: 'read', write: 'write' },
    beforeFirstWrite,
    masking: 'applied',
    ...overrides,
  })
  return {
    adapter,
    skill,
    read,
    executeTool,
    beforeFirstWrite,
    change: (patch: Partial<WorkspaceAdapterState>) => (state = { ...state, ...patch }),
    prepare: async (mode: 'edit' | 'reference' = 'edit', signal?: AbortSignal) =>
      adapter.prepare({ runId: 'run-P', state: await adapter.capture(), mode }, signal),
  }
}

describe('workspace document execution boundary', () => {
  it('exports templates outside prepared work and revalidates after an asynchronous export', async () => {
    const pending = deferred<void>()
    const snapshot = {
      fileName: 'draft.docx',
      title: 'draft',
      bytes: new Uint8Array([1, 2]),
      redacted: false,
    }
    const template = vi.fn(async () => {
      await pending.promise
      return snapshot
    })
    const f = fixture({ template })
    const state = await f.adapter.capture()
    const read = f.adapter.template!(state, 'preview')
    await vi.waitFor(() => expect(template).toHaveBeenCalledOnce())
    await expect(f.prepare()).rejects.toMatchObject({ kind: 'busy' })
    f.change({ revision: '2' })
    pending.resolve()
    await expect(read).rejects.toMatchObject({ kind: 'stale-document' })
    const prepared = await f.prepare()
    await expect(f.adapter.template!(prepared.state, 'info')).rejects.toMatchObject({
      kind: 'busy',
    })
    prepared.release()
    expect(await f.adapter.template!(await f.adapter.capture(), 'preview')).toEqual(snapshot)
    f.adapter.dispose()
  })

  it('rejects redacted as-is exports and discards cancelled template reads', async () => {
    const abort = new AbortController()
    const f = fixture({
      template: async () => ({
        fileName: 'draft.docx',
        title: 'draft',
        redacted: true,
        bytes: null,
      }),
    })
    const state = await f.adapter.capture()
    await expect(f.adapter.template!(state, 'as-is')).rejects.toMatchObject({
      kind: 'unsafe-content',
    })
    expect(await f.adapter.template!(state, 'info')).toMatchObject({ redacted: true, bytes: null })
    abort.abort()
    await expect(f.adapter.template!(state, 'preview', abort.signal)).rejects.toMatchObject({
      kind: 'cancelled',
    })
    f.adapter.dispose()
  })
  it('refreshes metadata without adopting a later caret or reading content', async () => {
    const f = fixture()
    const pinned = await f.adapter.capture()
    f.change({ selectionId: 'later-caret' })
    expect(await f.adapter.refresh(pinned)).toEqual(pinned)
    f.change({ revision: '2' })
    await expect(f.adapter.refresh(pinned)).rejects.toMatchObject({ kind: 'stale-selection' })
    const whole = { ...pinned, selectionId: null }
    expect(await f.adapter.refresh(whole)).toMatchObject({ revision: '2', selectionId: null })
    f.change({ instanceId: 'replacement' })
    await expect(f.adapter.refresh(whole)).rejects.toMatchObject({ kind: 'stale-document' })
    expect(f.read).not.toHaveBeenCalled()
    expect(f.executeTool).not.toHaveBeenCalled()
  })

  it('rejects a refresh hook that substitutes another selection', async () => {
    const f = fixture({ refreshSelection: () => 'another-selection' })
    await expect(f.adapter.refresh(await f.adapter.capture())).rejects.toMatchObject({
      kind: 'stale-selection',
    })
  })

  it('copies captured state, tool schemas and inputs across asynchronous work', async () => {
    const gate = deferred<void>()
    const beforeFirstWrite = vi.fn(() => gate.promise)
    const host = fixture({ beforeFirstWrite })
    const session = await host.prepare()
    const state = session.state
    expect(Object.isFrozen(state)).toBe(true)
    const tools = session.tools as AgentToolDef[]
    tools[0].name = 'forged'
    expect(session.tools[0].name).toBe('read')
    const request = call()
    const pending = session.execute(request)
    request.input.value = 'changed by composer'
    gate.resolve()
    await pending
    expect(host.executeTool.mock.calls[0][0].input.value).toBe('replacement')
    expect(state.revision).toBe('1')
    expect(session.state.revision).toBe('2')
  })

  it.each(['reference', 'readOnly'] as const)('offers only reading in %s mode', async (mode) => {
    const host = fixture()
    if (mode === 'readOnly') host.change({ readOnly: true })
    const session = await host.prepare(mode === 'reference' ? mode : 'edit')
    expect(session.tools.map((tool) => tool.name)).toEqual(['read'])
    await expect(session.execute(call())).rejects.toMatchObject({ kind: 'unknown-tool' })
    await session.execute(call('read'))
    expect(host.beforeFirstWrite).not.toHaveBeenCalled()
  })

  it('never promotes an unlisted tool or an inherited property into a capability', async () => {
    const host = fixture()
    const session = await host.prepare()
    for (const name of ['network', 'toString', '__proto__']) {
      await expect(session.execute(call(name))).rejects.toMatchObject({ kind: 'unknown-tool' })
    }
    expect(host.executeTool).not.toHaveBeenCalled()
  })

  it.each([
    { inputError: 'invalid JSON' },
    { truncated: true },
    { input: {} },
    { input: { value: null } },
    { input: { value: [] } },
    { input: [] },
    { id: '' },
  ])('rejects malformed tool calls before executing: %j', async (patch) => {
    const host = fixture()
    const session = await host.prepare()
    await expect(session.execute({ ...call(), ...patch } as AgentToolCall)).rejects.toMatchObject({
      kind: 'invalid-call',
    })
    expect(host.executeTool).not.toHaveBeenCalled()
    expect(host.beforeFirstWrite).not.toHaveBeenCalled()
  })

  it('rejects duplicate capabilities instead of choosing a schema by order', async () => {
    const host = fixture()
    host.skill.tools.push(definition('write'))
    await expect(host.prepare()).rejects.toMatchObject({ kind: 'invalid-call' })
  })

  it('does not follow the live caret and advances only its own document revision', async () => {
    const host = fixture()
    const session = await host.prepare()
    host.change({ selectionId: 'later-caret' })
    await session.execute(call('read', 'read-1'))
    await session.execute(call('write', 'write-1'))
    await session.execute(call('write', 'write-2'))
    expect(session.state).toMatchObject({ revision: '3', selectionId: 'captured-range' })
    expect(host.beforeFirstWrite).toHaveBeenCalledOnce()
    expect(await session.verifyResponse('claimed action', [{ name: 'read', ok: true }])).toBe(
      'correction',
    )
    expect(host.skill.verifyResponse).toHaveBeenCalledWith('claimed action', [
      { name: 'read', ok: true },
    ])
  })

  it.each([
    { instanceId: 'replacement-document' },
    { revision: 'manual-edit' },
    { readOnly: true },
  ])('rechecks state after the first-write hook: %j', async (patch) => {
    const gate = deferred<void>()
    const beforeFirstWrite = vi.fn(() => gate.promise)
    const host = fixture({ beforeFirstWrite })
    const session = await host.prepare()
    const pending = session.execute(call())
    await vi.waitFor(() => expect(beforeFirstWrite).toHaveBeenCalledOnce())
    host.change(patch)
    gate.resolve()
    await expect(pending).rejects.toMatchObject({
      kind: patch.readOnly ? 'read-only' : 'stale-document',
    })
    expect(host.executeTool).not.toHaveBeenCalled()
  })

  it('does not retry an already executed call id', async () => {
    const host = fixture()
    const session = await host.prepare()
    await session.execute(call())
    await expect(session.execute(call())).rejects.toMatchObject({ kind: 'duplicate-call' })
    expect(host.executeTool).toHaveBeenCalledOnce()
  })

  it('keeps the adapter busy until a released preparation actually settles', async () => {
    const gate = deferred<AgentSkill>()
    const host = fixture({ createSkill: () => gate.promise })
    const entered = vi.fn()
    host.skill.buildContext = entered
    const preparing = host.prepare()
    const rejected = expect(preparing).rejects.toMatchObject({ kind: 'disposed' })
    await Promise.resolve()
    host.adapter.dispose()
    gate.resolve(host.skill)
    await rejected
    expect(host.executeTool).not.toHaveBeenCalled()
  })

  it('rejects cancellation during preparation without retaining an active session', async () => {
    const gate = deferred<AgentSkill>()
    let attempts = 0
    const host = fixture({ createSkill: () => (++attempts === 1 ? gate.promise : host.skill) })
    const controller = new AbortController()
    const pending = host.prepare('edit', controller.signal)
    await vi.waitFor(() => expect(attempts).toBe(1))
    controller.abort()
    await expect(host.prepare()).rejects.toMatchObject({ kind: 'busy' })
    gate.resolve(host.skill)
    await expect(pending).rejects.toMatchObject({ kind: 'cancelled' })
    const next = await host.prepare()
    await next.execute(call())
  })

  it('aborts a released call and prevents overlapping document executions while it settles', async () => {
    const gate = deferred<void>()
    const started = deferred<AbortSignal | undefined>()
    const host = fixture()
    host.skill.executeTool = async (_call, signal) => {
      started.resolve(signal)
      await gate.promise
      return { output: 'late', summary: 'read' }
    }
    const session = await host.prepare()
    const pending = session.execute(call('read'))
    const signal = await started.promise
    await expect(session.execute(call('read', 'other'))).rejects.toMatchObject({ kind: 'busy' })
    session.release()
    expect(signal?.aborted).toBe(true)
    await expect(host.prepare()).rejects.toMatchObject({ kind: 'busy' })
    gate.resolve()
    await expect(pending).rejects.toMatchObject({ kind: 'cancelled' })
    await expect(host.prepare()).resolves.toMatchObject({ runId: 'run-P' })
    await expect(session.execute(call())).rejects.toMatchObject({ kind: 'released' })
  })

  it.each(['read-mutation', 'unreported-edit', 'unversioned-write'] as const)(
    'rejects inconsistent executor result: %s',
    async (mode) => {
      const host = fixture()
      host.skill.executeTool = () => {
        if (mode === 'unreported-edit') host.change({ revision: '2' })
        return { output: 'done', summary: 'fixture', mutated: mode !== 'unreported-edit' }
      }
      const session = await host.prepare()
      await expect(
        session.execute(call(mode === 'read-mutation' ? 'read' : 'write')),
      ).rejects.toMatchObject({ kind: 'stale-document' })
      await expect(session.execute(call())).rejects.toMatchObject({ kind: 'released' })
    },
  )

  it('rejects a removed capability even if it was offered during preparation', async () => {
    const host = fixture()
    const session = await host.prepare()
    host.skill.tools = host.skill.tools.filter((tool) => tool.name !== 'write')
    await expect(session.execute(call())).rejects.toMatchObject({ kind: 'unknown-tool' })
    expect(host.executeTool).not.toHaveBeenCalled()
  })
})

describe('current document reference snapshots', () => {
  it('returns only editor-approved text, the captured version and masking support', async () => {
    const host = fixture()
    const captured = await host.adapter.capture()
    const read = await host.adapter.read(captured, { scope: 'document', maxChars: 7 })
    expect(read).toEqual({
      text: 'current',
      truncated: true,
      masking: 'applied',
      scope: 'document',
      state: captured,
    })
    expect(host.read).toHaveBeenCalledWith(captured, { scope: 'document', maxChars: 7 }, undefined)
  })

  it('defaults to unsupported masking and caps oversized reads', async () => {
    const host = fixture({ masking: undefined })
    await expect(
      host.adapter.read(await host.adapter.capture(), {
        scope: 'document',
        maxChars: WORKSPACE_READ_MAX_CHARS + 1,
      }),
    ).resolves.toMatchObject({ masking: 'unsupported' })
    expect(host.read.mock.calls[0]).toEqual([
      expect.anything(),
      { scope: 'document', maxChars: WORKSPACE_READ_MAX_CHARS },
      undefined,
    ])
  })

  it.each([0, -1, Infinity, NaN, 1.5])('rejects an invalid read limit %s', async (maxChars) => {
    const host = fixture()
    await expect(
      host.adapter.read(await host.adapter.capture(), { scope: 'document', maxChars }),
    ).rejects.toMatchObject({ kind: 'unsupported-read' })
    expect(host.read).not.toHaveBeenCalled()
  })

  it('rejects stale, missing-selection and cancelled reads before touching the reader', async () => {
    const host = fixture()
    const captured = await host.adapter.capture()
    await expect(
      host.adapter.read({ ...captured, selectionId: null }, { scope: 'selection' }),
    ).rejects.toMatchObject({ kind: 'stale-selection' })
    const controller = new AbortController()
    controller.abort()
    await expect(
      host.adapter.read(captured, { scope: 'document' }, controller.signal),
    ).rejects.toMatchObject({ kind: 'cancelled' })
    host.change({ revision: '2' })
    await expect(host.adapter.read(captured, { scope: 'document' })).rejects.toMatchObject({
      kind: 'stale-document',
    })
    expect(host.read).not.toHaveBeenCalled()
  })

  it('discards a read if the document changes before it finishes', async () => {
    const gate = deferred<{ text: string; truncated: boolean }>()
    const started = vi.fn(() => gate.promise)
    const host = fixture({ read: started })
    const pending = host.adapter.read(await host.adapter.capture(), { scope: 'document' })
    await vi.waitFor(() => expect(started).toHaveBeenCalledOnce())
    host.change({ revision: 'manual-edit' })
    gate.resolve({ text: 'outdated', truncated: false })
    await expect(pending).rejects.toMatchObject({ kind: 'stale-document' })
  })
})
