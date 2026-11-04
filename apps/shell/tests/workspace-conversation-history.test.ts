import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ProjectStore, type ChatMessage } from '@genoffice/project-store'
import type { WorkspaceConversationOrigin } from '../src/shared/workspace-ai-api'
import { WorkspaceConversationHistory } from '../src/main/workspace-ai/conversation-history'

describe('workspace conversation history origins', () => {
  let root: string
  let local: ProjectStore
  const projects = {
    openChat: vi.fn(),
    append: vi.fn(),
    newChat: vi.fn(),
  }
  const sender = { editor: 'docs' as const, docPath: null }
  const memory = (id: string, windowSessionId = 'window-1'): WorkspaceConversationOrigin => ({
    kind: 'window-memory',
    conversationId: id,
    windowSessionId,
  })
  beforeEach(() => {
    vi.resetAllMocks()
    root = mkdtempSync(join(tmpdir(), 'history-origins-'))
    local = new ProjectStore(root)
  })
  afterEach(() => rmSync(root, { recursive: true, force: true }))

  it('opens and appends to legacy history by its stable reference after first save and rename', async () => {
    const path = join(root, 'first.docx')
    local.appendChatMessage('default', 'unsaved-one', { role: 'user', text: 'before save' })
    local.appendChatMessage('default', 'unsaved-one', { role: 'assistant', text: 'reply' })
    const ref = local.rebindChatToFile('default', 'unsaved-one', path)
    const origin: WorkspaceConversationOrigin = { kind: 'legacy-document', ...ref }
    local.fileRenamed(path, join(root, 'renamed.docx'))
    const history = new WorkspaceConversationHistory(local, projects, 'window-1')
    const file = join(root, 'projects', ref.projectId, 'chats', `${ref.chatId}.jsonl`)
    const before = readFileSync(file, 'utf8')
    expect((await history.load(origin)).messages.map((m) => m.text)).toEqual([
      'before save',
      'reply',
    ])
    expect(readFileSync(file, 'utf8')).toBe(before)
    await history.append(origin, { role: 'assistant', text: 'continued' }, sender)
    expect(local.loadChat(ref.projectId, ref.chatId).at(-1)?.text).toBe('continued')
    expect(local.listWorkspaceChats()).toEqual([])
    expect(projects.openChat).not.toHaveBeenCalled()
  })

  it('routes project records through the existing trust/foreign/write-guard service with no local fallback', async () => {
    const origin: WorkspaceConversationOrigin = {
      kind: 'project',
      projectId: 'project-one',
      chatId: 'chat-one',
    }
    const records = [
      {
        seq: 0,
        ts: '2026-11-04T00:00:00Z',
        role: 'assistant' as const,
        text: 'foreign',
        origin: 'other-pc',
        editor: 'docs',
      },
    ]
    projects.openChat.mockResolvedValue({ status: 'ok', messages: records, foreign: true })
    projects.append.mockResolvedValue({ status: 'unwritable' })
    const history = new WorkspaceConversationHistory(local, projects, 'window-1')
    expect(await history.load(origin)).toEqual({ messages: records, foreign: true })
    expect(await history.append(origin, { role: 'user', text: 'question' }, sender)).toBe(
      'unwritable',
    )
    expect(projects.append).toHaveBeenCalledWith(
      { projectId: 'project-one', chatId: 'chat-one', role: 'user', text: 'question' },
      sender,
    )
    projects.openChat.mockResolvedValue({ status: 'not-found' })
    await expect(history.load(origin)).rejects.toThrow(/not available/)
    expect(readdirSync(root)).toEqual([])
  })

  it('persists only new local conversations and leaves window-memory histories isolated', async () => {
    const history = new WorkspaceConversationHistory(local, projects, 'window-1')
    await history.append(memory('home-P'), { role: 'assistant', text: 'P memory' }, sender)
    expect((await history.load(memory('home-Q'))).messages).toEqual([])
    expect(readdirSync(root)).toEqual([])
    const origin = history.createLocal('New workspace conversation')
    await history.append(origin, { role: 'user', text: 'new question' }, sender)
    await history.append(origin, { role: 'assistant', text: 'new answer' }, sender)
    const restarted = new WorkspaceConversationHistory(new ProjectStore(root), projects, 'window-2')
    expect((await restarted.load(origin)).messages.map((m) => m.text)).toEqual([
      'new question',
      'new answer',
    ])
    await expect(restarted.load(memory('home-P'))).rejects.toThrow(/another window/)
    expect((await restarted.load(memory('home-P', 'window-2'))).messages).toEqual([])
    expect((await history.load(memory('home-P'))).messages[0]?.text).toBe('P memory')
  })

  it('keeps the home transcript at its original live source without creating local history', async () => {
    const messages: ChatMessage[] = [
      { role: 'assistant', text: 'home', seq: 0, ts: '2026-11-04T00:00:00Z' },
    ]
    const history = new WorkspaceConversationHistory(local, projects, 'window-1')
    history.attachMemory('home', {
      load: () => messages,
      append: (m) => {
        messages.push({ ...m, seq: messages.length, ts: '2026-11-04T00:00:00Z' })
      },
    })
    await history.append(memory('home'), { role: 'assistant', text: 'still home' }, sender)
    expect(messages.at(-1)?.text).toBe('still home')
    messages.push({ ...messages[0]!, text: 'new home event', seq: 2 })
    expect((await history.load(memory('home'))).messages.at(-1)?.text).toBe('new home event')
    expect(readdirSync(root)).toEqual([])
  })

  it('opens explicitly created project conversations before their first reply materializes files', async () => {
    projects.newChat.mockResolvedValue({ chatId: 'new-chat' })
    projects.openChat.mockResolvedValue({ status: 'not-found' })
    const history = new WorkspaceConversationHistory(local, projects, 'window-1')
    const origin = await history.createProject('project-one')
    expect((await history.load(origin)).messages).toEqual([])
    const pending = { role: 'user' as const, text: 'opening', seq: 0, ts: '2026-11-04T00:00:00Z' }
    projects.append.mockResolvedValue({ status: 'pending', message: pending })
    await history.append(origin, { role: 'user', text: 'opening' }, sender)
    expect((await history.load(origin)).messages).toEqual([pending])
    projects.append.mockResolvedValue({ status: 'unwritable' })
    await history.append(origin, { role: 'assistant', text: 'answer' }, sender)
    await expect(history.load(origin)).rejects.toThrow(/not available/)
    expect(readdirSync(root)).toEqual([])
  })
})
