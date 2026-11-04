import { describe, expect, it, vi } from 'vitest'
import { AgentLoop, type AgentStreamCallbacks, type AgentTransport } from '@genoffice/agent-core'
import type { ChatMessage } from '@genoffice/project-store'
import type {
  WorkspaceBoundDocument,
  WorkspaceConversationOrigin,
} from '../src/shared/workspace-ai-api'
import {
  WorkspaceConversations,
  type WorkspaceConversationRun,
} from '../src/renderer/src/workspace-ai/conversations'

const origin = (id: string): WorkspaceConversationOrigin => ({
  kind: 'workspace',
  conversationId: id,
})
const message = (text: string): ChatMessage => ({
  role: 'assistant',
  text,
  seq: 0,
  ts: '2026-11-04T00:00:00Z',
})
const target = (id: string): WorkspaceBoundDocument => ({
  document: { windowId: 1, tabId: id, documentId: id, kind: 'docs', generation: 1 },
  revision: 'v1',
  selection: { id: 'selection', revision: 'v1' },
})
function fixture() {
  const deps = {
    load: vi.fn(async () => ({ messages: [message('restored')], foreign: false })),
    createRuntime: vi.fn(() => ({ cancel: vi.fn(), dispose: vi.fn() })),
    changed: vi.fn(),
  }
  return { deps, controller: new WorkspaceConversations(deps) }
}

describe('workspace conversation isolation', () => {
  it('preserves draft, attachments, skills, targets, project, scroll and MCP choices across selection', async () => {
    const { controller: c, deps } = fixture()
    const p = await c.open(origin('P'), target('A'))
    const q = await c.open(origin('Q'), target('B'))
    const composer = {
      ...c.state(p).composer,
      text: 'draft',
      attachments: [{ name: 'one.pdf' }],
      skills: [{ id: 'local:s', source: 'local', title: 'Skill' }],
      project: { id: 'project-P', contextVersion: '1' },
      references: [{ ...target('C'), snapshotId: 'snapshot-C' }],
    }
    c.edit(p, {
      composer,
      scrollTop: 250,
      mcp: { disabledServers: ['server'], alwaysAllowedTools: ['tool'] },
    })
    c.select(p)
    c.select(q)
    c.select(p)
    expect(c.state(p).composer).toEqual(composer)
    expect(c.state(p).scrollTop).toBe(250)
    expect(c.state(q).composer.text).toBe('')
    expect(c.state(q).composer.project).toBeNull()
    expect(c.state(q).mcp.alwaysAllowedTools).toEqual([])
    expect(c.runtime(p)).not.toBe(c.runtime(q))
    expect(c.runtime(p).cancel).not.toHaveBeenCalled()
    const copy = c.state(p)
    copy.composer.text = 'outside'
    composer.attachments[0]!.name = 'outside'
    expect(c.state(p).composer.text).toBe('draft')
    expect(c.state(p).composer.attachments[0]?.name).toBe('one.pdf')
    await c.open(origin('P'), target('B'))
    expect(c.state(p).composer.target?.document.tabId).toBe('A')
    expect(deps.load).toHaveBeenCalledTimes(2)
  })

  it('routes delayed messages, errors and permissions to P while Q is selected', async () => {
    const { controller: c } = fixture()
    const p = await c.open(origin('P'))
    const q = await c.open(origin('Q'))
    const run = c.enqueue(p, 'run')
    expect(c.start(run)).toBe(true)
    const permission = c.ask(run, 'allow', { tool: 'write' })
    c.select(q)
    c.deliver(run, { message: message('P result') })
    c.deliver(run, { error: 'P error' })
    expect(c.state(p).messages.at(-1)?.text).toBe('P result')
    expect(c.state(p).error).toBe('P error')
    expect(c.state(p).permissions).toHaveLength(1)
    expect(c.state(q).messages.map((m) => m.text)).toEqual(['restored'])
    expect(c.state(q).permissions).toEqual([])
    expect(c.state(q).error).toBeNull()
    expect(c.answer({ ...run, conversationId: q }, 'allow', true)).toBe(false)
    expect(c.answer(run, 'allow', true)).toBe(true)
    await expect(permission).resolves.toBe(true)
    await expect(c.ask(run, 'allow', { tool: 'write' })).resolves.toBe(false)
    expect(c.runtime(p).cancel).not.toHaveBeenCalled()
  })

  it('freezes queued input and rejects callbacks from cancelled runs even after run-id reuse', async () => {
    const { controller: c } = fixture()
    const p = await c.open(origin('P'), target('A'))
    c.edit(p, { composer: { ...c.state(p).composer, text: 'first' } })
    const old = c.enqueue(p, 'same')
    c.start(old)
    c.edit(p, { composer: { ...c.state(p).composer, text: 'later', target: target('B') } })
    expect(c.state(p).runs[0]?.input.text).toBe('first')
    expect(c.state(p).runs[0]?.input.target?.document.tabId).toBe('A')
    const signal = c.signal(old)!
    const permission = c.ask(old, 'allow', {})
    c.finish(old, true)
    expect(signal.aborted).toBe(true)
    await expect(permission).resolves.toBe(false)
    const next = c.enqueue(p, 'same')
    expect(c.start(next)).toBe(true)
    expect(c.deliver(old, { message: message('stale') })).toBe(false)
    expect(c.finish(old, true)).toBe(false)
    expect(c.answer(old, 'allow', true)).toBe(false)
    expect(c.runtime(p).cancel).toHaveBeenCalledTimes(1)
    expect(c.deliver(next, { progress: 'new' })).toBe(true)
  })

  it('cancels only the specified queued request without cancelling the running conversation', async () => {
    const { controller: c } = fixture()
    const p = await c.open(origin('P'))
    const q = await c.open(origin('Q'))
    const a = c.enqueue(p, 'a')
    const b = c.enqueue(q, 'b')
    expect(c.start(a)).toBe(true)
    expect(c.start(b)).toBe(false)
    expect(c.deliver(b, { error: 'not started' })).toBe(false)
    c.finish(b, true)
    expect(c.signal(a)?.aborted).toBe(false)
    expect(c.runtime(p).cancel).not.toHaveBeenCalled()
    expect(c.runtime(q).cancel).not.toHaveBeenCalled()
  })

  it('deduplicates parallel opens and never revives a disposed controller after a slow load', async () => {
    const { controller: c, deps } = fixture()
    let resolve!: (value: { messages: ChatMessage[]; foreign: boolean }) => void
    deps.load.mockImplementation(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const first = c.open(origin('P'))
    const duplicate = c.open(origin('P'))
    expect(duplicate).toBe(first)
    await Promise.resolve()
    c.dispose()
    resolve({ messages: [message('late')], foreign: false })
    await expect(first).rejects.toThrow(/closed/)
    expect(deps.createRuntime).not.toHaveBeenCalled()
  })

  it('rejects the previous panel controller tickets even when conversation/run ids are reused', async () => {
    const { controller: first } = fixture()
    const p = await first.open(origin('P'))
    const old = first.enqueue(p, 'same')
    first.dispose()
    const { controller: next } = fixture()
    const reopened = await next.open(origin('P'))
    const current = next.enqueue(reopened, 'same')
    next.start(current)
    expect(old.generation).toBe(current.generation)
    expect(next.deliver(old, { error: 'old panel' })).toBe(false)
    expect(next.finish(old, true)).toBe(false)
    expect(next.signal(current)?.aborted).toBe(false)
  })

  it('follows an explicit first-save rebind while preserving the loop, draft, execution and permission', async () => {
    const { controller: c, deps } = fixture()
    const original: WorkspaceConversationOrigin = {
      kind: 'legacy-document',
      projectId: 'default',
      chatId: 'unsaved-one',
    }
    const id = await c.open(original)
    c.edit(id, { composer: { ...c.state(id).composer, text: 'unfinished draft' } })
    const runtime = c.runtime(id)
    const run = c.enqueue(id, 'run')
    c.start(run)
    const permission = c.ask(run, 'allow', {})
    const ref = { projectId: 'default', chatId: 'saved-chat-id' }
    c.rebindLegacy(id, ref)
    expect(await c.open({ kind: 'legacy-document', ...ref })).toBe(id)
    expect(await c.open(original)).toBe(id)
    expect(c.runtime(id)).toBe(runtime)
    expect(c.state(id).composer.text).toBe('unfinished draft')
    expect(c.state(id).origin).toEqual({ kind: 'legacy-document', ...ref })
    c.answer(run, 'allow', true)
    await expect(permission).resolves.toBe(true)
    expect(c.deliver(run, { message: message('still the same run') })).toBe(true)
    expect(deps.load).toHaveBeenCalledTimes(1)
    const other = await c.open({ kind: 'legacy-document', projectId: 'default', chatId: 'other' })
    expect(() => c.rebindLegacy(other, ref)).toThrow(/merge/)
  })

  it('denies outstanding permissions and cancels every runtime when the panel is disposed', async () => {
    const { controller: c } = fixture()
    const p = await c.open(origin('P'))
    const q = await c.open(origin('Q'))
    const a = c.runtime(p)
    const b = c.runtime(q)
    const run = c.enqueue(p, 'run')
    c.start(run)
    const permission = c.ask(run, 'allow', {})
    c.dispose()
    c.dispose()
    await expect(permission).resolves.toBe(false)
    expect(a.cancel).toHaveBeenCalledTimes(1)
    expect(b.dispose).toHaveBeenCalledTimes(1)
    expect(c.deliver(run, { message: message('late') })).toBe(false)
  })

  it('keeps real AgentLoop histories and delayed transport callbacks independent of the selected conversation', async () => {
    interface Runtime {
      loop: AgentLoop
      run: WorkspaceConversationRun | null
      cancel(): void
    }
    const callbacks = new Map<string, AgentStreamCallbacks>()
    const streams = vi.fn()
    const c: WorkspaceConversations<Runtime> = new WorkspaceConversations({
      load: async () => ({ messages: [], foreign: false }),
      createRuntime: (id) => {
        const transport: AgentTransport = {
          stream: (_request, cb) => {
            streams(id)
            callbacks.set(id, cb)
            return { cancel: vi.fn() }
          },
        }
        const runtime: Runtime = {
          run: null,
          cancel: () => runtime.loop.cancel(),
          loop: new AgentLoop({
            transport,
            skill: {
              id: 'conversation-test',
              systemPrompt: '',
              tools: [],
              buildContext: () => '',
              executeTool: () => ({ output: '', summary: '' }),
            },
            events: {
              onText: (text) => {
                if (runtime.run) c.deliver(runtime.run, { progress: text })
              },
              onDone: ({ text }) => {
                if (!runtime.run) return
                c.deliver(runtime.run, { message: message(text) })
                c.finish(runtime.run)
              },
            },
          }),
        }
        return runtime
      },
    })
    const p = await c.open(origin('P'))
    const q = await c.open(origin('Q'))
    c.select(p)
    expect(streams).not.toHaveBeenCalled()
    const runP = c.enqueue(p, 'P')
    c.start(runP)
    c.runtime(p).run = runP
    await c.runtime(p).loop.run('question P')
    c.select(q)
    callbacks.get(p)!.onDelta('answer P')
    callbacks.get(p)!.onDone()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(c.state(p).messages.at(-1)?.text).toBe('answer P')
    expect(c.state(q).messages).toEqual([])
    expect(c.runtime(q).loop.messages).toEqual([])
    const runQ = c.enqueue(q, 'Q')
    c.start(runQ)
    c.runtime(q).run = runQ
    await c.runtime(q).loop.run('question Q')
    callbacks.get(q)!.onDelta('answer Q')
    callbacks.get(q)!.onDone()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(
      c
        .runtime(p)
        .loop.messages.map((m) => (m.role === 'tool' ? '' : m.text))
        .join(' '),
    ).toContain('answer P')
    expect(
      c
        .runtime(q)
        .loop.messages.map((m) => (m.role === 'tool' ? '' : m.text))
        .join(' '),
    ).not.toContain('question P')
    expect(streams).toHaveBeenCalledTimes(2)
    c.dispose()
  })
})
