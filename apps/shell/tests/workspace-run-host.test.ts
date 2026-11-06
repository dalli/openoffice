import { afterEach, describe, expect, it, vi } from 'vitest'
import { makeWorkspaceHost, until, tick, deferred } from './helpers/workspace-host-fixture'
import type { ChatMessage } from '@genoffice/project-store'
import type { WorkspacePendingSend } from '../src/renderer/src/workspace-ai/run-host'

const cleanups: Array<() => void> = []
afterEach(() => {
  for (const off of cleanups.splice(0)) off()
})
function fixture(...args: Parameters<typeof makeWorkspaceHost>) {
  const f = makeWorkspaceHost(...args)
  cleanups.push(f.dispose)
  return f
}

describe('workspace model host through the real execution boundary', () => {
  it('captures project context only after the gate admits, and binds the held request to its original conversation', async () => {
    let held: WorkspacePendingSend | null = null
    const captureProject = vi.fn(async () => ({ id: 'project', contextVersion: 'prepared' }))
    let ready = false
    const f = fixture({
      admit: (request) => {
        if (!ready) held = request
        return ready
      },
      captureProject,
      prepare: async (input) => ({ project: input.project, skills: [], dispose() {} }),
    })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'A')
    expect(await p.send()).toBeNull()
    expect(captureProject).not.toHaveBeenCalled()
    f.host.conversations.select(q.id)
    ready = true
    const run = (await f.host.resume(held!))!
    await until(() => f.requests.length === 1)
    expect(captureProject).toHaveBeenCalledWith(expect.objectContaining({ conversationId: p.id }))
    expect(f.host.conversations.state(p.id).runs[0]!.input.project).toEqual({
      id: 'project',
      contextVersion: 'prepared',
    })
    expect(f.requests[0]!.conversation).toBe(p.id)
    expect(f.host.conversations.selectedId).toBe(q.id)
    await f.host.cancel(run.run)
  })

  it('cancels a project admission before its async context read returns and discards the late capability', async () => {
    const reading = deferred()
    const captureProject = vi.fn(async () => {
      await reading.promise
      return { id: 'project', contextVersion: 'late' }
    })
    const discardProject = vi.fn(async () => {})
    const f = fixture({ captureProject, discardProject })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const submitted = p.send()
    await until(() => captureProject.mock.calls.length === 1)
    const run = f.host.conversations.state(p.id).runs[0]!.run
    await f.host.cancel({ ...run, generation: run.generation + 1 })
    expect(f.host.conversations.state(p.id).runs).toHaveLength(1)
    await f.host.cancel(structuredClone(run))
    reading.resolve()
    expect(await submitted).toBeNull()
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(p.id).runs).toEqual([])
    expect(discardProject).toHaveBeenCalledWith(p.id, { id: 'project', contextVersion: 'late' })
  })
  it('refreshes only the original idle transcript and keeps its draft, runtime, selection and MCP choices', async () => {
    let records: ChatMessage[] = []
    const f = fixture({
      load: async () => ({ messages: structuredClone(records), foreign: false }),
    })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'A')
    const runtime = f.host.conversations.runtime(p.id)
    f.host.conversations.edit(p.id, {
      scrollTop: 33,
      mcp: { disabledServers: ['one'], alwaysAllowedTools: [] },
    })
    f.host.conversations.select(q.id)
    records = [
      { seq: 0, ts: new Date().toISOString(), role: 'assistant', text: 'external original reply' },
    ]
    f.host.historyChanged(f.host.conversations.state(p.id).origin)
    await until(() => f.host.conversations.state(p.id).messages.length === 1)
    expect(f.host.conversations.runtime(p.id)).toBe(runtime)
    expect(runtime.loop.messages).toEqual([{ role: 'assistant', text: 'external original reply' }])
    expect(f.host.conversations.selectedId).toBe(q.id)
    expect(f.host.conversations.state(q.id).messages).toEqual([])
    expect(f.host.conversations.state(p.id)).toMatchObject({
      scrollTop: 33,
      composer: { text: 'edit A' },
      mcp: { disabledServers: ['one'] },
    })
    expect(f.requests).toEqual([])
  })

  it('does not overwrite an in-flight transcript with a late history read', async () => {
    const waiting = deferred()
    let delay = false
    const f = fixture({
      load: async () => {
        if (delay) await waiting.promise
        return {
          messages: [
            { seq: 0, ts: new Date().toISOString(), role: 'assistant' as const, text: 'external' },
          ],
          foreign: false,
        }
      },
    })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    delay = true
    f.host.historyChanged(f.host.conversations.state(p.id).origin)
    expect(f.host.conversations.state(p.id).messages.at(-1)?.role).toBe('user')
    f.callbacks.get(p.id)![0]!.onDelta('in-flight answer')
    f.callbacks.get(p.id)![0]!.onDone()
    await run.done
    expect(f.host.conversations.state(p.id).messages.at(-1)?.text).toBe('in-flight answer')
    waiting.resolve()
    await until(() => f.host.conversations.state(p.id).messages.at(-1)?.text === 'external')
    expect(f.requests).toHaveLength(1)
  })

  it('refuses a first-save collision instead of merging two open runtimes', async () => {
    const f = fixture()
    f.addEditor('A')
    const before = { kind: 'legacy-document' as const, projectId: 'default', chatId: 'unsaved-one' }
    const after = { ...before, chatId: 'saved-one' }
    const target = await f.execution.capture('A')
    const p = await f.host.conversations.open(before, target.binding)
    const q = await f.host.conversations.open(after, target.binding)
    const original = f.host.conversations.runtime(p)
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'draft' },
    })
    f.host.rebindHistory(before, after)
    expect(f.host.conversations.runtime(p)).toBe(original)
    expect(f.host.conversations.runtime(p)).not.toBe(f.host.conversations.runtime(q))
    expect(f.host.conversations.state(p).origin).toEqual(before)
    expect(f.host.conversations.state(p).error).toContain('Cannot merge')
    expect(await f.host.submit(p, target.id, [])).toBeNull()
    expect(f.requests).toEqual([])
  })

  it('accepts cloned UI run handles and rejects another controller, conversation, run or generation', async () => {
    const f = fixture()
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const submitted = (await p.send())!
    await until(() => f.requests.length === 1)
    const copied = f.host.conversations.state(p.id).runs[0]!.run
    expect(copied).toEqual(submitted.run)
    expect(copied).not.toBe(submitted.run)
    for (const invalid of [
      { ...copied, controllerId: 'another-controller' },
      { ...copied, conversationId: 'another-conversation' },
      { ...copied, runId: 'another-run' },
      { ...copied, generation: copied.generation + 1 },
    ])
      await f.host.cancel(invalid)
    expect(f.cancellations).not.toHaveBeenCalled()
    await f.host.cancel(copied)
    await submitted.done
    expect(f.cancellations).toHaveBeenCalledTimes(1)
  })
  it('uses the updated stable record source if first save rebinds between the two history writes', async () => {
    const origins: unknown[] = []
    let id = ''
    const f = fixture({
      append: async (origin) => {
        origins.push(origin)
        if (origins.length === 1)
          f.host.conversations.rebindLegacy(id, { projectId: 'default', chatId: 'saved-chat' })
        return 'recorded'
      },
    })
    f.addEditor('A')
    const target = await f.execution.capture('A')
    id = await f.host.conversations.open(
      { kind: 'legacy-document', projectId: 'default', chatId: 'unsaved-chat' },
      target.binding,
    )
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'question' },
    })
    const submitted = (await f.host.submit(id, target.id, []))!
    await until(() => f.requests.length === 1)
    f.callbacks.get(id)![0]!.onDelta('answer')
    f.callbacks.get(id)![0]!.onDone()
    await submitted.done
    expect(origins).toEqual([
      { kind: 'legacy-document', projectId: 'default', chatId: 'unsaved-chat' },
      { kind: 'legacy-document', projectId: 'default', chatId: 'saved-chat' },
    ])
  })
  it('resumes the exact request held by the shared send gate without following another draft or selected conversation', async () => {
    let held: WorkspacePendingSend | null = null
    let allowed = false
    const f = fixture({
      admit: (request) => {
        if (!allowed) held = request
        return allowed
      },
    })
    f.addEditor('A')
    f.addEditor('B')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    expect(await p.send()).toBeNull()
    f.host.conversations.edit(p.id, {
      composer: {
        ...f.host.conversations.state(p.id).composer,
        text: 'later unsent draft',
        target: q.target.binding,
      },
    })
    f.host.conversations.select(q.id)
    allowed = true
    const submitted = (await f.host.resume(held!))!
    await until(() => f.requests.length === 1)
    expect(f.requests[0]?.conversation).toBe(p.id)
    expect(
      f.requests[0]?.request.messages.some(
        (m) =>
          m.role === 'user' && m.text.includes('edit A') && !m.text.includes('later unsent draft'),
      ),
    ).toBe(true)
    expect(f.host.conversations.state(p.id).composer.text).toBe('later unsent draft')
    f.callbacks.get(p.id)![0]!.onDelta('original response')
    f.callbacks.get(p.id)![0]!.onDone()
    await submitted.done
    expect(f.host.conversations.state(q.id).messages).toEqual([])
  })
  it('uses the explicit unsaved reference and edits only the target while another conversation is selected; undo restores the whole run', async () => {
    const f = fixture()
    const source = f.addEditor('A')
    const target = f.addEditor('B')
    f.addEditor('C')
    const p = await f.conversation('P', 'B', ['A'])
    const q = await f.conversation('Q', 'C')
    f.host.conversations.select(p.id)
    expect(f.requests).toEqual([])
    const submitted = (await p.send())!
    await until(() => f.requests.length === 1)
    f.host.conversations.select(q.id)
    expect(
      f.requests[0]!.request.messages.some(
        (m) => m.role === 'user' && m.text.includes('A UNSAVED'),
      ),
    ).toBe(true)
    const first = f.callbacks.get(p.id)![0]!
    first.onToolCall({ id: 'write-one', name: 'write', input: { text: 'B uses A UNSAVED' } })
    first.onDone()
    await until(() => f.callbacks.get(p.id)!.length === 2)
    const second = f.callbacks.get(p.id)![1]!
    second.onToolCall({ id: 'write-two', name: 'write', input: { text: 'B final from A UNSAVED' } })
    second.onDone()
    await until(() => f.callbacks.get(p.id)!.length === 3)
    f.callbacks.get(p.id)![2]!.onDelta('completed B')
    f.callbacks.get(p.id)![2]!.onDone()
    await submitted.done
    expect(target.body()).toBe('B final from A UNSAVED')
    expect(source.body()).toBe('A UNSAVED')
    expect(f.host.conversations.state(q.id).messages).toEqual([])
    expect(f.host.conversations.state(p.id).messages.at(-1)?.text).toBe('completed B')
    expect(f.append).toHaveBeenCalledTimes(2)
    const card = f.host.rollbackCards(p.id)[0]!
    expect(card.state).toBe('applied')
    expect(f.host.rollbackCards(q.id)).toEqual([])
    await f.host.rollback(q.id, card.id)
    expect(target.body()).toBe('B final from A UNSAVED')
    await f.host.rollback(p.id, card.id)
    expect(target.body()).toBe('B UNSAVED')
    expect(f.host.rollbackCards(p.id)[0]?.state).toBe('reverted')
  })

  it('freezes queued drafts and cancels a queued request in the same conversation without stopping the running loop', async () => {
    const f = fixture()
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const active = (await p.send())!
    await until(() => f.requests.length === 1)
    f.host.conversations.edit(p.id, {
      composer: { ...f.host.conversations.state(p.id).composer, text: 'queued draft' },
    })
    const queued = (await p.send())!
    await f.host.cancel(queued.run)
    await queued.done
    expect(f.cancellations).not.toHaveBeenCalled()
    f.callbacks.get(p.id)![0]!.onDelta('active reply')
    f.callbacks.get(p.id)![0]!.onDone()
    await active.done
    expect(f.append).toHaveBeenCalledTimes(2)
    expect(f.requests).toHaveLength(1)
  })

  it('waits for cancelled editor work to settle before starting the next conversation', async () => {
    const f = fixture()
    const a = f.addEditor('A')
    f.addEditor('B')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    a.gate.write = deferred()
    const active = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks
      .get(p.id)![0]!
      .onToolCall({ id: 'late-write', name: 'write', input: { text: 'must never apply' } })
    f.callbacks.get(p.id)![0]!.onDone()
    await tick()
    const queued = (await q.send())!
    await f.host.cancel(active.run)
    await tick()
    expect(f.requests).toHaveLength(1)
    expect(f.cancellations).toHaveBeenCalled()
    a.gate.write.resolve()
    await until(() => f.requests.length === 2)
    expect(a.body()).toBe('A UNSAVED')
    f.callbacks.get(q.id)![0]!.onDelta('B answer')
    f.callbacks.get(q.id)![0]!.onDone()
    await queued.done
  })

  it('waits for transport cancellation before reusing the same conversation loop', async () => {
    const f = fixture()
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    const second = (await p.send())!
    await f.host.cancel(first.run)
    await until(() => f.requests.length === 2)
    f.callbacks.get(p.id)![1]!.onDelta('second run')
    f.callbacks.get(p.id)![1]!.onDone()
    await second.done
    expect(f.host.conversations.state(p.id).messages.at(-1)?.text).toBe('second run')
  })

  it('denies an owned permission before waiting for a blocked tool to finish cancelling', async () => {
    const f = fixture({
      prepare: async (_input, _run, permissions) => ({
        project: null,
        dispose() {},
        skills: [
          {
            id: 'approval',
            systemPrompt: '',
            tools: [{ name: 'ask', description: 'Ask', inputSchema: { type: 'object' } }],
            executeTool: async () => ({
              output: await permissions.ask({
                server: 'fixture',
                tool: 'ask',
                exposedName: 'ask',
                input: {},
                description: '',
                destructive: false,
              }),
              summary: 'answered',
            }),
          },
        ],
      }),
    })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const submitted = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks.get(p.id)![0]!.onToolCall({ id: 'permission', name: 'ask', input: {} })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.host.permissions(p.id).requests().length === 1)
    await f.host.cancel(submitted.run)
    expect(f.host.permissions(p.id).requests()).toEqual([])
    expect(f.requests).toHaveLength(1)
  })

  it('stops on reference changes and refuses an old undo token after manual edits', async () => {
    const f = fixture()
    const a = f.addEditor('A')
    const b = f.addEditor('B')
    const p = await f.conversation('P', 'B', ['A'])
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks.get(p.id)![0]!.onToolCall({ id: 'one', name: 'write', input: { text: 'AI B' } })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.callbacks.get(p.id)!.length === 2)
    a.manual()
    await run.done
    expect(f.cancellations).toHaveBeenCalled()
    const card = f.host.rollbackCards(p.id)[0]!
    b.manual()
    await f.host.rollback(p.id, card.id)
    expect(b.body()).toBe('B MANUAL')
    expect(f.host.rollbackCards(p.id)[0]?.state).toBe('refused')
  })

  it('keeps a rejected send in its composer and rechecks the shared gate before a queued model call', async () => {
    const f = fixture()
    f.addEditor('A')
    f.addEditor('B')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.setGate({ kind: 'checking' })
    expect(await p.send()).toBeNull()
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(p.id).composer.text).toBe('edit A')
    f.setGate({ kind: 'run', choice: { connected: true, provider: 'claude-code', metered: false } })
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    const queued = (await q.send())!
    f.setGate({ kind: 'checking' })
    f.callbacks.get(p.id)![0]!.onDelta('reply')
    f.callbacks.get(p.id)![0]!.onDone()
    await first.done
    await tick()
    expect(f.requests).toHaveLength(1)
    f.setGate({ kind: 'run', choice: { connected: true, provider: 'claude-code', metered: false } })
    f.host.wake()
    await until(() => f.requests.length === 2)
    f.callbacks.get(q.id)![0]!.onDelta('queued reply')
    f.callbacks.get(q.id)![0]!.onDone()
    await queued.done
    expect(f.requests).toHaveLength(2)
    expect(f.host.conversations.state(q.id).composer.text).toBe('edit B')
  })

  it('keeps the reply visible and reports a project write failure without retry or local fallback', async () => {
    const append = vi.fn(async () => 'unwritable' as const)
    const f = fixture({ append })
    f.addEditor('A')
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks.get(p.id)![0]!.onDelta('answer')
    f.callbacks.get(p.id)![0]!.onDone()
    await run.done
    expect(f.host.conversations.state(p.id).messages.at(-1)?.text).toBe('answer')
    expect(f.host.conversations.state(p.id).error).toBe('history-unwritable')
    expect(append).toHaveBeenCalledTimes(1)
    expect(f.requests).toHaveLength(1)
  })
})

describe('workspace asynchronous checks before each model turn', () => {
  it('cancels a pending check without starting the transport or accepting a late answer', async () => {
    const wait = deferred()
    const beforeTurn = vi.fn(() => wait.promise)
    const f = makeWorkspaceHost({
      prepare: async () => ({ project: null, skills: [], beforeTurn, dispose() {} }),
    })
    try {
      f.addEditor('A')
      const p = await f.conversation('P', 'A')
      const sent = (await p.send())!
      await until(() => beforeTurn.mock.calls.length === 1)
      await f.host.cancel(sent.run)
      wait.resolve()
      await sent.done
      expect(f.requests).toEqual([])
      expect(f.host.busy()).toBe(false)
    } finally {
      f.dispose()
    }
  })
  it('checks the next turn again and stops without sending when consent was withdrawn', async () => {
    let allowed = true
    const beforeTurn = vi.fn(async () => {
      if (!allowed) throw new Error('Drive consent changed')
    })
    const f = makeWorkspaceHost({
      prepare: async () => ({ project: null, skills: [], beforeTurn, dispose() {} }),
    })
    try {
      f.addEditor('A')
      const p = await f.conversation('P', 'A')
      const sent = (await p.send())!
      await until(() => f.requests.length === 1)
      allowed = false
      const cb = f.callbacks.get(p.id)![0]!
      cb.onToolCall({ id: 'read', name: 'read', input: {} })
      cb.onDone()
      await sent.done
      expect(beforeTurn).toHaveBeenCalledTimes(2)
      expect(f.requests).toHaveLength(1)
    } finally {
      f.dispose()
    }
  })
})
