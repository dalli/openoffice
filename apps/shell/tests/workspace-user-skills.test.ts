import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  WorkspaceUserSkills,
  type WorkspaceSkillsBridge,
} from '../src/renderer/src/workspace-ai/user-skills'
import { makeWorkspaceHost, deferred, until } from './helpers/workspace-host-fixture'
import type { WorkspaceComposer } from '../src/renderer/src/workspace-ai/conversations'
import type { SkillTemplateBridge } from '@genoffice/ui'

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach((off) => off()))
const entry = (name = 'one', extra = {}) => ({
  id: `app:${name}`,
  name,
  source: 'app',
  title: name,
  description: `Use ${name}`,
  enabled: true,
  invoke: 'auto',
  owned: true,
  ...extra,
})
const chip = (name = 'one') => ({ id: `app:${name}`, source: 'app', title: name })
function fixture(
  patch: Partial<Parameters<typeof makeWorkspaceHost>[0]> = {},
  template?: SkillTemplateBridge,
  mcpServerUsable?: (conversationId: string, server: string) => boolean,
  readTemplate?: ConstructorParameters<typeof WorkspaceUserSkills>[5],
) {
  let items = [entry()]
  let body = 'PRIVATE_SKILL_BODY !`never-execute-this`'
  let changed = () => {}
  const bridge: WorkspaceSkillsBridge = {
    list: vi.fn(async () => items),
    aiRead: vi.fn(async (id) =>
      items.some((item) => item.id === id && item.enabled)
        ? {
            ok: true,
            skill: {
              ...items.find((item) => item.id === id),
              body,
              files: ['notes.md'],
              mcpServers: ['absent'],
              hasScripts: true,
            },
          }
        : { ok: false, reason: 'disabled' },
    ),
    aiReadFile: vi.fn(async () => ({ ok: true, text: 'PRIVATE_SKILL_FILE' })),
    openSettings: vi.fn(async () => {}),
    onChanged: vi.fn((listener) => {
      changed = listener
      return vi.fn()
    }),
  }
  const f = makeWorkspaceHost({ prepare: (input, run) => skills.prepare(input, run), ...patch })
  const skills = new WorkspaceUserSkills(
    bridge,
    f.host,
    () => 'en',
    template ?? null,
    mcpServerUsable,
    readTemplate,
  )
  f.addEditor('A')
  f.addEditor('B')
  cleanups.push(() => {
    skills.dispose()
    f.dispose()
  })
  const draft = (id: string, update: Partial<WorkspaceComposer>) =>
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, ...update },
    })
  return {
    ...f,
    skills,
    bridge,
    draft,
    setItems: (next: typeof items) => {
      items = next
      changed()
    },
    setBody: (next: string) => {
      body = next
    },
  }
}

describe('workspace user skills through the existing skill and execution guards', () => {
  it('runs the copied template with current skill text once, without opening another copy', async () => {
    const openTemplate = vi.fn()
    const f = fixture({}, { openTemplate } as unknown as SkillTemplateBridge)
    f.setItems([entry('one', { template: 'assets/form.docx', templateExists: true })])
    const p = await f.conversation('P', 'A')
    f.skills.acceptHandoff(p.id, await f.execution.capture('A'), {
      kind: 'template-handoff',
      id: 'one',
      text: 'Fill',
      chips: [chip()],
    })
    f.draft(p.id, { skills: [chip()] })
    f.setBody('CURRENT_TEMPLATE_SKILL_BODY')
    let run: Awaited<ReturnType<typeof p.send>>
    await f.skills.send(p.id, async () => {
      run = await p.send()
    })
    await until(() => f.requests.length === 1)
    expect(openTemplate).not.toHaveBeenCalled()
    expect(JSON.stringify(f.requests[0]!.request)).toContain('CURRENT_TEMPLATE_SKILL_BODY')
    f.callbacks.get(p.id)![0]!.onDelta('Done')
    f.callbacks.get(p.id)![0]!.onDone()
    await run!.done
    expect(JSON.stringify(f.append.mock.calls)).not.toContain('CURRENT_TEMPLATE_SKILL_BODY')
    expect(f.host.conversations.state(p.id).messages[0]!.skills).toEqual([chip()])
    const send = vi.fn(async () => {})
    f.draft(p.id, { target: (await f.execution.capture('B')).binding, skills: [chip()] })
    await f.skills.send(p.id, send)
    expect(send).not.toHaveBeenCalled()
    expect(openTemplate).not.toHaveBeenCalled()
  })

  it('does not treat handed chips in another conversation as copied templates', async () => {
    const f = fixture()
    f.setItems([entry('one', { template: 'assets/form.docx', templateExists: true })])
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.skills.acceptHandoff(p.id, await f.execution.capture('A'), {
      kind: 'template-handoff',
      id: 'one',
      text: 'Fill',
      chips: [chip()],
    })
    f.draft(q.id, { skills: [chip()] })
    const run = (await q.send())!
    await run.done
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(q.id).error).not.toBeNull()
  })

  it('keeps template actions and preview completion on the proposing conversation and original target', async () => {
    let ended: (raw: unknown) => void = () => {}
    const off = vi.fn()
    const template = {
      onPreviewEnded: vi.fn((listener) => {
        ended = listener
        return off
      }),
    } as unknown as SkillTemplateBridge
    const read = vi.fn(async () => ({
      fileName: 'A.docx',
      title: 'A',
      redacted: false,
      bytes: new Uint8Array([1, 2]),
    }))
    const f = fixture({}, template, undefined, read)
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks.get(p.id)![0]!.onToolCall({
      id: 'proposal',
      name: 'propose_skill',
      input: {
        name: 'template-one',
        description: 'Reusable',
        body: 'Use this form',
        wantsTemplate: true,
      },
    })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.requests.length === 2)
    f.callbacks.get(p.id)![1]!.onDelta('Done')
    f.callbacks.get(p.id)![1]!.onDone()
    await run.done
    const card = f.skills.proposals(p.id)[0]!
    f.host.conversations.select(q.id)
    const document = f.skills.document(card)!
    expect(await document.redacted()).toBe(false)
    expect(await document.bytes({ allowRedacted: false })).toEqual(new Uint8Array([1, 2]))
    expect(read.mock.calls).toEqual([
      [expect.objectContaining({ tabId: 'A' }), 'info'],
      [expect.objectContaining({ tabId: 'A' }), 'as-is'],
    ])
    expect(document.fileName).toBe('A.docx')
    read.mockResolvedValueOnce({
      fileName: 'A.docx',
      title: 'A',
      redacted: true,
      bytes: new Uint8Array([3]),
    })
    await document.bytes()
    read.mockResolvedValueOnce({
      fileName: 'renamed.odt',
      title: 'renamed',
      redacted: false,
      bytes: new Uint8Array([]),
    })
    expect(await document.redacted()).toBe(true)
    expect(document.fileName).toBe('A.docx')
    f.skills.previewOpened(card, { tabId: 'copy', previewId: 'preview' })
    ended({ previewId: 'other', outcome: 'saved', skillId: 'app:wrong' })
    expect(card.status).toBe('preview')
    ended({
      previewId: 'preview',
      outcome: 'saved',
      skillId: 'app:template-one',
      title: 'template-one',
    })
    await until(() => card.status === 'saved')
    expect(f.host.conversations.state(q.id).messages).toEqual([])
    expect(f.append.mock.calls.at(-1)![0]).toEqual(f.host.conversations.state(p.id).origin)
    expect(JSON.stringify(f.append.mock.calls)).not.toContain('A.docx')
    expect(f.requests).toHaveLength(2)
  })
  it('uses the original conversation MCP choices for selected and automatic skill instructions', async () => {
    let usable = true
    const f = fixture({}, undefined, (id, server) => id === p.id && server === 'absent' && usable)
    f.setItems([entry(), entry('auto')])
    const p = await f.conversation('P', 'A')
    f.draft(p.id, { skills: [chip()] })
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    expect(JSON.stringify(f.requests[0]!.request)).not.toContain('cannot be used right now')
    usable = false
    f.callbacks
      .get(p.id)![0]!
      .onToolCall({ id: 'use', name: 'use_skill', input: { id: 'app:auto' } })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.requests.length === 2)
    expect(JSON.stringify(f.requests[1]!.request)).toContain(
      'MCP server absent, which cannot be used right now',
    )
    f.callbacks.get(p.id)![1]!.onDelta('Done')
    f.callbacks.get(p.id)![1]!.onDone()
    await run.done
  })

  it('adds selected instructions only to the original model request and keeps bodies out of history', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.draft(p.id, { skills: [chip()] })
    const run = (await p.send())!
    f.host.conversations.select(q.id)
    await until(() => f.requests.length === 1)
    expect(f.requests[0]!.conversation).toBe(p.id)
    expect(JSON.stringify(f.requests[0]!.request)).toContain('PRIVATE_SKILL_BODY')
    expect(JSON.stringify(f.requests[0]!.request)).toContain('never-execute-this')
    expect(JSON.stringify(f.requests[0]!.request)).toContain('absent')
    f.callbacks.get(p.id)![0]!.onDelta('Finished')
    f.callbacks.get(p.id)![0]!.onDone()
    await run.done
    expect(JSON.stringify(f.append.mock.calls)).not.toContain('PRIVATE_SKILL_BODY')
    expect(f.append.mock.calls[0]![1]).toMatchObject({ text: 'edit A', skills: [chip()] })
    expect(f.host.conversations.state(p.id).messages[0]!.text).toBe('edit A')
    expect(f.host.conversations.state(q.id).messages).toEqual([])
  })

  it('rereads a queued skill and refuses one turned off while another conversation runs', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    f.draft(q.id, { skills: [chip()], text: 'Queued original draft' })
    const second = (await q.send())!
    // Match the panel's successful enqueue clearing its own composer.
    f.draft(q.id, { skills: [], text: '' })
    f.setItems([entry('one', { enabled: false })])
    f.callbacks.get(p.id)![0]!.onDone()
    await Promise.all([first.done, second.done])
    expect(f.requests).toHaveLength(1)
    expect(f.host.conversations.state(q.id).composer).toMatchObject({
      text: 'Queued original draft',
      skills: [chip()],
    })
    expect(f.skills.composer(q.id).notice).toContain('nothing was sent')
  })

  it('blocks oversized instructions and preserves the chosen chips without a model call', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    f.setBody('x'.repeat(150_001))
    f.draft(p.id, { skills: [chip()] })
    await (await p.send())!.done
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(p.id).composer.skills).toEqual([chip()])
    expect(f.skills.composer(p.id).notice).toBeTruthy()
  })

  it('discards a late skill read after cancellation and never sends into the newly selected conversation', async () => {
    const f = fixture()
    const waiting = deferred()
    vi.mocked(f.bridge.aiRead).mockImplementation(async () => {
      await waiting.promise
      return { ok: true, skill: { ...entry(), body: 'LATE_BODY', files: [] } }
    })
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.draft(p.id, { skills: [chip()] })
    const run = (await p.send())!
    await until(() => vi.mocked(f.bridge.aiRead).mock.calls.length === 1)
    f.host.conversations.select(q.id)
    await f.host.cancel(run.run)
    waiting.resolve()
    await run.done
    expect(f.requests).toEqual([])
    expect(f.skills.proposals(q.id)).toEqual([])
  })

  it('uses the existing auto tools, omits skill/file bodies from records, and routes proposals to their original conversation', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    expect(JSON.stringify(f.requests[0]!.request)).toContain('app:one')
    f.host.conversations.select(q.id)
    f.callbacks
      .get(p.id)![0]!
      .onToolCall({ id: 'use', name: 'use_skill', input: { id: 'app:one' } })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.requests.length === 2)
    expect(JSON.stringify(f.requests[1]!.request)).toContain('PRIVATE_SKILL_BODY')
    f.callbacks.get(p.id)![1]!.onToolCall({
      id: 'read',
      name: 'read_skill_file',
      input: { id: 'app:one', path: 'notes.md' },
    })
    f.callbacks.get(p.id)![1]!.onDone()
    await until(() => f.requests.length === 3)
    expect(JSON.stringify(f.requests[2]!.request)).toContain('PRIVATE_SKILL_FILE')
    f.callbacks.get(p.id)![2]!.onToolCall({
      id: 'proposal',
      name: 'propose_skill',
      input: { name: 'new-one', description: 'New', body: 'PROPOSED_PRIVATE_BODY' },
    })
    f.callbacks.get(p.id)![2]!.onDone()
    await until(() => f.requests.length === 4)
    expect(f.skills.proposals(p.id)[0]!.proposal.name).toBe('new-one')
    expect(f.skills.proposals(q.id)).toEqual([])
    f.callbacks.get(p.id)![3]!.onDelta('Done')
    f.callbacks.get(p.id)![3]!.onDone()
    await run.done
    expect(JSON.stringify(f.append.mock.calls)).not.toMatch(/PRIVATE_SKILL|PROPOSED_PRIVATE/)
    expect(f.append.mock.calls[1]![1]!.tools).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ name: 'use_skill', summary: expect.any(String) }),
        expect.objectContaining({
          name: 'propose_skill',
          output: expect.stringContaining('new-one'),
        }),
      ]),
    )
  })

  it('restores historical chips once after opening, rereads current text and skips an unavailable historical skill', async () => {
    const f = fixture({
      load: async () => ({
        foreign: false,
        messages: [
          {
            role: 'user',
            seq: 0,
            ts: new Date().toISOString(),
            text: 'Old request',
            skills: [chip(), chip('gone')],
          },
        ],
      }),
    })
    const p = await f.conversation('P', 'A')
    f.setBody('CURRENT_BODY')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    expect(JSON.stringify(f.requests[0]!.request)).toContain('CURRENT_BODY')
    f.callbacks.get(p.id)![0]!.onDone()
    await run.done
    const reads = vi.mocked(f.bridge.aiRead).mock.calls.length
    const next = (await p.send())!
    await until(() => f.requests.length === 2)
    expect(vi.mocked(f.bridge.aiRead).mock.calls.length).toBe(reads)
    f.callbacks.get(p.id)![1]!.onDone()
    await next.done
  })

  it('hands a selected template to the existing copy flow and keeps a newer draft and the other conversation', async () => {
    const waiting = deferred()
    const openTemplate = vi.fn(async () => {
      await waiting.promise
      return { ok: true, tabId: 'template-copy' }
    })
    const template = { openTemplate } as unknown as SkillTemplateBridge
    const f = fixture({}, template)
    f.setItems([entry('one', { template: 'assets/form.docx', templateExists: true })])
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.draft(p.id, { skills: [chip()] })
    const send = vi.fn(async () => {})
    const sending = f.skills.send(p.id, send)
    await until(() => openTemplate.mock.calls.length === 1)
    f.host.conversations.select(q.id)
    f.draft(p.id, { text: 'Newer draft' })
    waiting.resolve()
    await sending
    expect(send).not.toHaveBeenCalled()
    expect(f.requests).toEqual([])
    expect(f.host.conversations.state(p.id).composer.text).toBe('Newer draft')
    expect(f.skills.templateNotes(p.id)[0]!.tabId).toBe('template-copy')
    expect(f.skills.templateNotes(q.id)).toEqual([])
  })

  it('rejects a path outside a skill and ignores a delayed automatic read after stop', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.callbacks.get(p.id)![0]!.onToolCall({
      id: 'outside',
      name: 'read_skill_file',
      input: { id: 'app:one', path: '../secret.txt' },
    })
    f.callbacks.get(p.id)![0]!.onDone()
    await until(() => f.requests.length === 2)
    expect(f.bridge.aiReadFile).not.toHaveBeenCalled()
    const waiting = deferred()
    vi.mocked(f.bridge.aiRead).mockImplementation(async () => {
      await waiting.promise
      return {
        ok: true,
        skill: { ...entry(), body: 'LATE', files: [], template: 'assets/template.docx' },
      }
    })
    f.callbacks
      .get(p.id)![1]!
      .onToolCall({ id: 'late', name: 'use_skill', input: { id: 'app:one' } })
    f.callbacks.get(p.id)![1]!.onDone()
    await until(() => vi.mocked(f.bridge.aiRead).mock.calls.length === 1)
    const stopped = f.host.cancel(run.run)
    waiting.resolve()
    await stopped
    await run.done
    expect(f.skills.templateNotes(p.id)).toEqual([])
    expect(f.requests).toHaveLength(2)
  })
})
