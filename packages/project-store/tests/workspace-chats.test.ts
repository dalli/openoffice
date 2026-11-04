import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ProjectStore } from '../src/store.js'

describe('local workspace history', () => {
  let root: string
  let store: ProjectStore
  const path = (id: string, ext: string) =>
    join(root, 'projects', '.workspace', 'v1', `${id}.${ext}`)
  beforeEach(() => {
    root = mkdtempSync(join(tmpdir(), 'workspace-history-'))
    store = new ProjectStore(root)
  })
  afterEach(() => {
    vi.restoreAllMocks()
    rmSync(root, { recursive: true, force: true })
  })

  it('keeps random-id history across restart with the existing record limits', () => {
    const a = store.createWorkspaceChat('A'.repeat(300))
    const b = store.createWorkspaceChat('B')
    expect(a.id).not.toBe(b.id)
    expect(a.title).toHaveLength(200)
    store.appendWorkspaceChatMessage(a.id, {
      role: 'user',
      text: 'question',
      skills: [{ id: 'local:one', source: 'local', title: 'One' }],
    })
    expect(existsSync(path(a.id, 'jsonl'))).toBe(false)
    expect(store.loadWorkspaceChat(a.id).map((m) => m.text)).toEqual(['question'])
    store.appendWorkspaceChatMessage(a.id, {
      role: 'assistant',
      text: 'x'.repeat(40_000),
      tools: [{ name: 'read', summary: 'Read', output: 'y'.repeat(20_000) }],
    })
    const reopened = new ProjectStore(root)
    const records = reopened.loadWorkspaceChat(a.id)
    expect(records.map((m) => m.seq)).toEqual([0, 1])
    expect(records[0]?.skills?.[0]?.id).toBe('local:one')
    expect(records[1]?.text).toHaveLength(32_000 + '\n\n[truncated]'.length)
    expect(records[1]?.tools?.[0]?.output).toHaveLength(16_000)
    reopened.appendWorkspaceChatMessage(a.id, { role: 'assistant', text: 'next' })
    expect(reopened.loadWorkspaceChat(a.id, 1)[0]?.seq).toBe(2)
    expect(reopened.loadWorkspaceChat(a.id, 0)).toHaveLength(1)
    reopened.renameWorkspaceChat(a.id, 'Renamed')
    expect(new ProjectStore(root).listWorkspaceChats().find((c) => c.id === a.id)?.title).toBe(
      'Renamed',
    )
  })

  it('never enrolls workspace conversations in document paths or legacy project lists', () => {
    const legacy = store.resolveChatForFile(join(root, 'document.docx'))
    store.appendChatMessage(legacy.projectId, legacy.chatId, { role: 'assistant', text: 'legacy' })
    const index = readFileSync(join(root, 'projects', 'index.json'), 'utf8')
    const info = store.createWorkspaceChat()
    store.appendWorkspaceChatMessage(info.id, { role: 'assistant', text: 'workspace' })
    // Even a legacy project's name and chat id matching the new namespace do not share counters.
    store.appendChatMessage('workspace', info.id, { role: 'assistant', text: 'separate' })
    expect(store.loadChat('workspace', info.id)[0]?.seq).toBe(0)
    expect(store.loadWorkspaceChat(info.id)[0]?.seq).toBe(0)
    expect(store.listChats(legacy.projectId).map((c) => c.chatId)).toEqual([legacy.chatId])
    expect(store.loadChat(legacy.projectId, legacy.chatId)[0]?.text).toBe('legacy')
    expect(readFileSync(join(root, 'projects', 'index.json'), 'utf8')).toBe(index)
  })

  it.each(['../escape', '/tmp/escape', 'a\\b', ''])('rejects unsafe id %j before IO', (id) => {
    expect(() => store.workspaceChatInfo(id)).toThrow()
    expect(() => store.loadWorkspaceChat(id)).toThrow()
    expect(() => store.appendWorkspaceChatMessage(id, { role: 'assistant', text: 'x' })).toThrow()
    expect(() => store.renameWorkspaceChat(id, 'title')).toThrow()
    expect(existsSync(join(root, 'projects'))).toBe(false)
  })

  it('preserves unsupported, malformed and oversized metadata without silently replacing it', () => {
    const info = store.createWorkspaceChat('original')
    for (const raw of [JSON.stringify({ ...info, version: 2 }), '{broken', 'x'.repeat(9_000)]) {
      writeFileSync(path(info.id, 'json'), raw)
      expect(store.listWorkspaceChats()).toEqual([])
      expect(() => store.loadWorkspaceChat(info.id)).toThrow(/unsupported/)
      expect(() =>
        store.appendWorkspaceChatMessage(info.id, { role: 'assistant', text: 'x' }),
      ).toThrow(/unsupported/)
      expect(() => store.renameWorkspaceChat(info.id, 'new')).toThrow(/unsupported/)
      expect(readFileSync(path(info.id, 'json'), 'utf8')).toBe(raw)
    }
  })

  it('continues after a corrupt JSONL tail using the last valid sequence', () => {
    const { id } = store.createWorkspaceChat()
    store.appendWorkspaceChatMessage(id, { role: 'assistant', text: 'valid' })
    writeFileSync(path(id, 'jsonl'), `${readFileSync(path(id, 'jsonl'), 'utf8')}{bad tail`)
    store = new ProjectStore(root)
    store.appendWorkspaceChatMessage(id, { role: 'assistant', text: 'continued' })
    expect(store.loadWorkspaceChat(id).map((m) => [m.seq, m.text])).toEqual([
      [0, 'valid'],
      [1, 'continued'],
    ])
  })

  it('surfaces write failures and retains the opening user message in memory', () => {
    const { id } = store.createWorkspaceChat()
    store.appendWorkspaceChatMessage(id, { role: 'user', text: 'pending' })
    mkdirSync(path(id, 'jsonl'))
    vi.spyOn(console, 'warn').mockImplementation(() => {})
    expect(() =>
      store.appendWorkspaceChatMessage(id, { role: 'assistant', text: 'answer' }),
    ).toThrow()
    expect(store.loadWorkspaceChat(id).map((m) => m.text)).toEqual(['pending'])
  })
})
