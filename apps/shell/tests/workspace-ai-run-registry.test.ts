import { describe, expect, it, vi } from 'vitest'
import { WorkspaceAiRunRegistry } from '../src/main/workspace-ai/run-registry'
import type {
  WorkspaceAiRunRequest,
  WorkspaceAiToolReply,
  WorkspaceAiToolRequest,
  WorkspaceDocumentEndpoint,
  WorkspaceDocumentIdentity,
} from '../src/shared/workspace-ai-api'

function setup(terminalLimit = 128) {
  let now = 1_000
  const owner = { windowId: 1, panelContentsId: 10, panelGeneration: 1 }
  const stopped = vi.fn()
  const registry = new WorkspaceAiRunRegistry(owner, () => now, 100, 32, stopped, terminalLimit)
  const document = (tabId: string): WorkspaceDocumentIdentity => ({
    windowId: 1,
    tabId,
    documentId: `identity-${tabId}`,
    kind: 'docs',
    generation: 1,
  })
  const a = document('a')
  const b = document('b')
  const c = document('c')
  const endpoint = (
    doc: WorkspaceDocumentIdentity,
    contentsId: number,
  ): WorkspaceDocumentEndpoint => ({
    document: doc,
    contentsId,
    revision: 'v1',
    readOnly: false,
    tools: [
      { name: 'read_document', access: 'read' },
      { name: 'replace_selection', access: 'write' },
    ],
  })
  registry.registerDocument(endpoint(a, 11))
  registry.registerDocument(endpoint(b, 12))
  registry.registerDocument(endpoint(c, 13))
  const run = (runId = 'run-p', conversationId = 'p'): WorkspaceAiRunRequest => ({
    owner,
    runId,
    conversationId,
    target: { document: b, revision: 'v1', selection: { id: 'selected-table', revision: 'v1' } },
    references: [{ document: a, revision: 'v1', selection: null, snapshotId: 'current-unsaved-a' }],
    project: { id: 'project-p', contextVersion: 'prepared-1' },
  })
  const call = (patch: Partial<WorkspaceAiToolRequest> = {}): WorkspaceAiToolRequest => ({
    owner,
    runId: 'run-p',
    bindingId: 1,
    conversationId: 'p',
    callId: 'call-1',
    document: b,
    tool: 'replace_selection',
    ...patch,
  })
  const reply = (patch: Partial<WorkspaceAiToolReply> = {}): WorkspaceAiToolReply => ({
    owner,
    runId: 'run-p',
    bindingId: 1,
    conversationId: 'p',
    callId: 'call-1',
    document: b,
    tool: 'replace_selection',
    revision: 'v2',
    mutated: true,
    ...patch,
  })
  const start = () => {
    expect(registry.enqueue(run()).ok).toBe(true)
    expect(registry.startNext()?.runId).toBe('run-p')
  }
  return {
    registry,
    owner,
    a,
    b,
    c,
    endpoint,
    run,
    call,
    reply,
    start,
    stopped,
    tick: () => (now += 100),
  }
}

describe('workspace AI run registry', () => {
  it('pins document, selection, project and conversation when the composer changes or another run is selected', () => {
    const { registry, run, call, b } = setup()
    const request = run()
    expect(registry.enqueue(request).ok).toBe(true)
    // A composer/tab transition mutates its own values, not the queued request.
    ;(request as { conversationId: string }).conversationId = 'q'
    ;(request.target.document as { documentId: string }).documentId = 'different-document'
    ;(request.project as { id: string }).id = 'project-q'
    const started = registry.startNext()!
    expect(started.conversationId).toBe('p')
    expect(started.target.document.documentId).toBe('identity-b')
    expect(started.target.selection?.id).toBe('selected-table')
    expect(started.project?.id).toBe('project-p')
    // The original caller's identity was changed, so use the pinned identity.
    expect(registry.claimTool(call({ document: { ...b, documentId: 'identity-b' } })).ok).toBe(true)
    expect(registry.claimTool(call({ conversationId: 'q', callId: 'wrong-chat' }))).toEqual({
      ok: false,
      reason: 'run-not-active',
    })
  })

  it('allows one active run per window, FIFO conversation queues and conversation-specific cancellation', () => {
    const { registry, owner, run } = setup()
    registry.enqueue(run())
    registry.enqueue(run('run-q', 'q'))
    registry.enqueue(run('run-p2', 'p'))
    expect(registry.startNext()?.runId).toBe('run-p')
    expect(registry.startNext()).toBeNull()
    expect(registry.cancel(owner, 'q', 'run-p', 1).ok).toBe(false)
    expect(registry.state('run-p')).toEqual({ status: 'running' })
    expect(registry.complete(owner, 'p', 'run-p', 1).ok).toBe(true)
    expect(registry.startNext()?.runId).toBe('run-q')
    expect(registry.cancel(owner, 'q', 'run-q', 2).ok).toBe(true)
    expect(registry.startNext()?.runId).toBe('run-p2')
  })

  it('rejects a foreign panel, another window document and reused run ids', () => {
    const { registry, owner, run, call, start, b } = setup()
    expect(registry.enqueue({ ...run(), owner: { ...owner, panelContentsId: 99 } })).toEqual({
      ok: false,
      reason: 'wrong-owner',
    })
    expect(
      registry.enqueue({
        ...run(),
        target: { ...run().target, document: { ...b, windowId: 2 } },
      }),
    ).toEqual({ ok: false, reason: 'missing-document' })
    start()
    expect(registry.enqueue(run())).toEqual({ ok: false, reason: 'duplicate-run' })
    expect(registry.claimTool(call({ owner: { ...owner, panelGeneration: 2 } }))).toEqual({
      ok: false,
      reason: 'wrong-owner',
    })
  })

  it('reads the explicit current reference but never promotes it to a write target or accesses a third tab', () => {
    const { registry, start, call, a, c } = setup()
    start()
    const reference = registry.claimTool(call({ document: a, tool: 'read_document' }))
    expect(reference).toMatchObject({
      ok: true,
      value: { contentsId: 11, access: 'read', revision: 'v1', snapshotId: 'current-unsaved-a' },
    })
    expect(
      registry.acceptReply(11, {
        ...call({ document: a, tool: 'read_document' }),
        revision: 'v1',
        mutated: false,
      }).ok,
    ).toBe(true)
    expect(registry.claimTool(call({ document: a, callId: 'write-a' }))).toEqual({
      ok: false,
      reason: 'read-only',
    })
    expect(
      registry.claimTool(call({ document: c, callId: 'read-c', tool: 'read_document' })),
    ).toEqual({
      ok: false,
      reason: 'unbound-document',
    })
    expect(registry.claimTool(call({ callId: 'shell', tool: 'execute_javascript' }))).toEqual({
      ok: false,
      reason: 'unknown-tool',
    })
  })

  it('requires the precise sender, generation, conversation, call and tool before accepting a reply', () => {
    const { registry, start, call, reply, b } = setup()
    start()
    expect(registry.claimTool(call()).ok).toBe(true)
    const wrong = [
      [13, reply()],
      [12, reply({ conversationId: 'q' })],
      [12, reply({ owner: { windowId: 1, panelContentsId: 10, panelGeneration: 2 } })],
      [12, reply({ bindingId: 99 })],
      [12, reply({ document: { ...b, generation: 2 } })],
      [12, reply({ callId: 'another-call' })],
      [12, reply({ tool: 'read_document' })],
    ] as const
    for (const [sender, response] of wrong) {
      expect(registry.acceptReply(sender, response)).toEqual({
        ok: false,
        reason: 'wrong-response',
      })
    }
    expect(registry.acceptReply(12, reply()).ok).toBe(true)
    expect(registry.acceptReply(12, reply())).toEqual({ ok: false, reason: 'wrong-response' })
  })

  it('consumes a call once, serializes calls and advances only its own verified mutations', () => {
    const { registry, start, call, reply } = setup()
    start()
    expect(registry.claimTool(call()).ok).toBe(true)
    expect(registry.claimTool(call())).toEqual({ ok: false, reason: 'duplicate-call' })
    expect(registry.claimTool(call({ callId: 'call-2' }))).toEqual({
      ok: false,
      reason: 'call-in-flight',
    })
    expect(registry.acceptReply(12, reply()).ok).toBe(true)
    expect(registry.claimTool(call())).toEqual({ ok: false, reason: 'duplicate-call' })
    expect(registry.claimTool(call({ callId: 'call-2' }))).toMatchObject({
      ok: true,
      value: { revision: 'v2' },
    })
  })

  it('refuses writes to a read-only target without disabling reads', () => {
    const { registry, b, endpoint, start, call } = setup()
    registry.registerDocument({ ...endpoint(b, 12), readOnly: true })
    start()
    expect(registry.claimTool(call())).toEqual({ ok: false, reason: 'read-only' })
    expect(registry.claimTool(call({ tool: 'read_document' })).ok).toBe(true)
  })

  it.each(['target', 'reference'] as const)(
    'invalidates %s changes before applying a late reply',
    (which) => {
      const { registry, start, call, reply, a, b } = setup()
      start()
      registry.claimTool(call())
      registry.documentChanged(which === 'target' ? b : a, 'changed-by-user')
      expect(registry.state('run-p')).toEqual({ status: 'stopped', reason: 'document-changed' })
      expect(registry.acceptReply(12, reply()).ok).toBe(false)
    },
  )

  it('invalidates old generations even if the tab id and contents id are reused', () => {
    const { registry, start, call, reply, b, endpoint } = setup()
    start()
    registry.claimTool(call())
    registry.registerDocument(endpoint({ ...b, generation: 2 }, 12))
    expect(registry.state('run-p')).toEqual({ status: 'stopped', reason: 'editor-replaced' })
    expect(registry.acceptReply(12, reply()).ok).toBe(false)
  })

  it('cancels closed documents without applying their result to a newly opened tab', () => {
    const { registry, start, call, reply, b, endpoint } = setup()
    start()
    registry.claimTool(call())
    registry.closeDocument(b)
    registry.registerDocument(endpoint({ ...b, documentId: 'new-document', generation: 2 }, 14))
    expect(registry.state('run-p')).toEqual({ status: 'stopped', reason: 'document-closed' })
    expect(registry.acceptReply(12, reply()).ok).toBe(false)
  })

  it('rejects timed-out or cancelled late replies and never resends an edit', () => {
    const { registry, start, call, reply, tick } = setup()
    start()
    registry.claimTool(call())
    tick()
    expect(registry.acceptReply(12, reply())).toEqual({ ok: false, reason: 'run-finished' })
    expect(registry.state('run-p')).toEqual({ status: 'stopped', reason: 'timeout' })
    expect(registry.claimTool(call())).toEqual({ ok: false, reason: 'timeout' })
    expect(registry.startNext()).toBeNull()
  })

  it('stops queued stale selections after an earlier run changed their target', () => {
    const { registry, start, run, call, reply } = setup()
    start()
    registry.enqueue(run('run-q', 'q'))
    registry.claimTool(call())
    registry.acceptReply(12, reply())
    expect(registry.state('run-q')).toEqual({ status: 'stopped', reason: 'document-changed' })
    expect(registry.startNext()).toBeNull()
  })

  it('requires read replies to preserve content and version', () => {
    const { registry, start, call, reply, a } = setup()
    start()
    registry.claimTool(call({ document: a, tool: 'read_document' }))
    expect(registry.acceptReply(11, reply({ document: a, tool: 'read_document' }))).toEqual({
      ok: false,
      reason: 'document-changed',
    })
  })

  it('keeps cancellation scoped to the run and drops every binding on panel close', () => {
    const { registry, start, owner, call, run, reply } = setup()
    start()
    registry.claimTool(call())
    registry.enqueue(run('run-q', 'q'))
    expect(registry.cancel(owner, 'q', 'run-q', 2).ok).toBe(true)
    expect(registry.state('run-p')).toEqual({ status: 'running' })
    registry.dispose()
    expect(registry.state('run-p')).toEqual({ status: 'stopped', reason: 'panel-closed' })
    expect(registry.acceptReply(12, reply()).ok).toBe(false)
    expect(registry.startNext()).toBeNull()
  })

  it.each(['state', 'claim'] as const)(
    'notifies transport cancellation even when %s discovers a timeout',
    (method) => {
      const { registry, start, call, tick, stopped } = setup()
      start()
      tick()
      if (method === 'state') registry.state('run-p')
      else registry.claimTool(call())
      expect(stopped).toHaveBeenCalledExactlyOnceWith({
        conversationId: 'p',
        runId: 'run-p',
        bindingId: 1,
        reason: 'timeout',
      })
      registry.expire()
      registry.state('run-p')
      expect(stopped).toHaveBeenCalledTimes(1)
    },
  )

  it('bounds terminal records without accepting an old reply or cancellation when a run id is reused', () => {
    const { registry, owner, run, call, reply } = setup(3)
    for (let bindingId = 1; bindingId <= 4; bindingId++) {
      const request = run(bindingId === 1 ? 'run-p' : `run-${bindingId}`)
      registry.enqueue(request)
      expect(registry.startNext()?.bindingId).toBe(bindingId)
      expect(registry.complete(owner, 'p', request.runId, bindingId).ok).toBe(true)
    }
    expect(registry.state('run-p')).toBeNull()
    expect(registry.state('run-2')).toEqual({ status: 'completed' })
    registry.enqueue(run())
    expect(registry.startNext()?.bindingId).toBe(5)
    expect(registry.claimTool(call()).ok).toBe(false)
    expect(registry.cancel(owner, 'p', 'run-p', 1).ok).toBe(false)
    expect(registry.complete(owner, 'p', 'run-p', 1).ok).toBe(false)
    expect(registry.claimTool(call({ bindingId: 5 })).ok).toBe(true)
    expect(registry.acceptReply(12, reply())).toEqual({ ok: false, reason: 'wrong-response' })
    expect(registry.acceptReply(12, reply({ bindingId: 5 })).ok).toBe(true)
  })

  it('rejects unknown access modes at registration rather than accidentally treating them as reads', () => {
    const { registry, b, endpoint } = setup()
    expect(() =>
      registry.registerDocument({
        ...endpoint(b, 12),
        tools: [{ name: 'dangerous', access: 'unknown' as 'read' }],
      }),
    ).toThrow('invalid workspace document endpoint')
  })

  it('invalidates an in-flight write when the target becomes read-only or its capability list changes', () => {
    const { registry, start, call, reply, b, endpoint, stopped } = setup()
    start()
    registry.claimTool(call())
    registry.registerDocument({ ...endpoint(b, 12), readOnly: true })
    expect(registry.acceptReply(12, reply()).ok).toBe(false)
    expect(stopped).toHaveBeenCalledExactlyOnceWith({
      conversationId: 'p',
      runId: 'run-p',
      bindingId: 1,
      reason: 'editor-replaced',
    })
  })
})
