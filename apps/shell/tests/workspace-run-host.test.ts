import { afterEach, describe, expect, it, vi } from 'vitest'
import { makeWorkspaceHost, until, tick, deferred } from './helpers/workspace-host-fixture'
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
