import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ProjectTrustAnswer } from '@genoffice/project-store'
import { ProjectGuide } from '../src/main/projects/project-guide'
import type { ProjectFolderServices } from '../src/main/projects/project-ipc'
import { WorkspaceProjectPreparation } from '../src/main/workspace-ai/project-preparation'
import type { WorkspaceAiBoundRun } from '../src/shared/workspace-ai-api'
import { prepareWorkspaceProject } from '../src/renderer/src/workspace-ai/project-preparation'
import { deferred } from './helpers/workspace-host-fixture'

const cleanups: Array<() => void> = []
afterEach(() => {
  for (const clean of cleanups.splice(0).reverse()) clean()
})

async function fixture(answer: ProjectTrustAnswer = 'use') {
  const base = mkdtempSync(join(tmpdir(), 'workspace-project-'))
  const root = join(base, 'project')
  const data = join(base, 'data')
  mkdirSync(root)
  mkdirSync(data)
  writeFileSync(join(root, 'AGENTS.md'), 'TRUSTED PROJECT RULES')
  const guide = new ProjectGuide({
    settingsPath: join(data, 'app-settings.json'),
    userData: data,
    homeDir: join(base, 'home'),
    defaultParentDir: () => base,
    env: {},
  })
  const trust = vi.fn(async () => answer)
  const added = await guide.addFolder(root, trust)
  if (added.status !== 'added') throw new Error('Project fixture was not registered')
  const opened = new Set<string>()
  const tools = {
    openScope: vi.fn(async () => {
      const id = `scope-${opened.size + 1}`
      opened.add(id)
      return id
    }),
    closeScope: vi.fn((_owner: number, raw: unknown) => {
      opened.delete((raw as { scope: string }).scope)
    }),
    listFiles: vi.fn(async () => ({ files: [], total: 0, offset: 0 })),
    search: vi.fn(async () => ({ matches: [], total: 0, offset: 0 })),
    read: vi.fn(async () => ({
      status: 'ok',
      text: 'REDACTED PROJECT MATERIAL',
      totalChars: 25,
      offset: 0,
    })),
  }
  const scopes = { projectOf: vi.fn(() => added.project.id) }
  const services = { guide, tools, scopes } as unknown as Pick<
    ProjectFolderServices,
    'guide' | 'tools' | 'scopes'
  >
  const memoryChanged = vi.fn()
  const service = new WorkspaceProjectPreparation(services, 10, memoryChanged)
  cleanups.push(() => {
    rmSync(base, { recursive: true, force: true })
  })
  cleanups.push(() => {
    service.dispose()
    guide.flushTouches()
  })
  const prepared = await service.prepare('conversation', added.project.id)
  const run: WorkspaceAiBoundRun = {
    owner: { windowId: 1, panelContentsId: 10, panelGeneration: 1 },
    conversationId: 'conversation',
    runId: 'run',
    bindingId: 1,
    project: prepared.binding,
    target: {
      document: { windowId: 1, tabId: 'A', documentId: 'doc', kind: 'docs', generation: 1 },
      revision: '1',
      selection: null,
    },
    references: [],
  }
  return {
    root,
    data,
    guide,
    trust,
    service,
    prepared,
    run,
    tools,
    opened,
    services,
    memoryChanged,
  }
}

describe('main-owned workspace project preparation', () => {
  it('announces a successful memory write through the existing project listeners', async () => {
    const f = await fixture()
    f.service.bind(f.run)
    await f.service.start(f.run)
    const result = await f.service.tool(f.run, 'remember', {
      change: { op: 'add', text: 'Useful project preference (2026-11-05)' },
    })
    expect(result).toMatchObject({ status: 'ok' })
    expect(f.memoryChanged).toHaveBeenCalledExactlyOnceWith(f.prepared.binding.id)
    expect(readFileSync(join(f.root, '.ieumoffice', 'memory.md'), 'utf8')).toContain(
      'Useful project preference',
    )
  })
  it('uses existing trust decisions and keeps context private until the bound run starts', async () => {
    const f = await fixture()
    expect(f.trust).toHaveBeenCalledOnce()
    expect(JSON.stringify(f.prepared)).not.toContain('TRUSTED PROJECT RULES')
    expect(f.tools.openScope).not.toHaveBeenCalled()
    expect(readdirSync(f.root)).toEqual(['AGENTS.md'])
    f.service.bind(f.run)
    expect(() => f.service.context(f.run)).toThrow()
    await f.service.start(f.run)
    expect(f.service.context(f.run).context.instructions?.text).toContain('TRUSTED PROJECT RULES')
    const copy = f.service.context(f.run)
    copy.context.instructions!.text = 'tampered'
    expect(f.service.context(f.run).context.instructions?.text).not.toBe('tampered')
    expect(() =>
      f.service.context({ ...f.run, owner: { ...f.run.owner, panelGeneration: 2 } }),
    ).toThrow()
    await f.service.tool(f.run, 'read', { path: 'material.md' })
    expect(f.tools.read).toHaveBeenCalledWith(10, { path: 'material.md', scope: 'scope-1' })
    f.service.release(f.run)
    expect(f.opened.size).toBe(0)
    await expect(f.service.tool(f.run, 'read', { path: 'material.md' })).rejects.toThrow()
    expect(readFileSync(join(f.root, 'AGENTS.md'), 'utf8')).toBe('TRUSTED PROJECT RULES')
  })

  it('honors use without instructions and does not expose the memory tool', async () => {
    const f = await fixture('without')
    f.service.bind(f.run)
    await f.service.start(f.run)
    const context = f.service.context(f.run)
    expect(context.context.instructions).toBeNull()
    expect(context.context.memory).toBeNull()
    expect(context.context.canRemember).toBe(false)
    await expect(
      f.service.tool(f.run, 'remember', { change: { op: 'add', text: 'untrusted' } }),
    ).rejects.toThrow()
    const prepared = await prepareWorkspaceProject(
      {
        projectContext: async () => context,
        projectTool: vi.fn(),
      },
      { binding: f.run, signal: new AbortController().signal } as never,
      () => 'en',
    )
    expect(prepared.skills[0]!.tools.map((t) => t.name)).toEqual([
      'project_list_files',
      'project_search',
      'project_read',
    ])
    expect(prepared.skills[0]!.systemPrompt).not.toContain('TRUSTED PROJECT RULES')
    prepared.dispose()
  })

  it('rejects a changed project between admission and start, including trust changes', async () => {
    const f = await fixture()
    f.service.bind(f.run)
    writeFileSync(join(f.root, 'AGENTS.md'), 'CHANGED WHILE QUEUED')
    await expect(f.service.start(f.run)).rejects.toThrow()
    expect(f.tools.openScope).not.toHaveBeenCalled()
    f.service.release(f.run)
    const next = await f.service.prepare('conversation', f.run.project!.id)
    const run = { ...f.run, bindingId: 2, runId: 'next', project: next.binding }
    f.service.bind(run)
    await f.guide.decideTrust(f.run.project!.id, 'without')
    await expect(f.service.start(run)).rejects.toThrow()
    expect(f.tools.openScope).not.toHaveBeenCalled()
  })

  it('rejects replay, another conversation and an unknown or removed project', async () => {
    const f = await fixture()
    expect(() => f.service.bind({ ...f.run, conversationId: 'other' })).toThrow()
    expect(() =>
      f.service.bind({ ...f.run, project: { ...f.prepared.binding, contextVersion: 'forged' } }),
    ).toThrow()
    f.service.bind(f.run)
    expect(() => f.service.bind({ ...f.run, bindingId: 2, runId: 'replay' })).toThrow()
    await expect(f.service.prepare('conversation', 'unknown')).rejects.toThrow()
    await f.guide.removeFromList(f.run.project!.id)
    await expect(f.service.start(f.run)).rejects.toThrow()
  })

  it('reclaims a scope that finishes opening after cancellation or panel disposal', async () => {
    const f = await fixture()
    f.service.bind(f.run)
    const opening = deferred()
    f.tools.openScope.mockImplementation(async () => {
      await opening.promise
      f.opened.add('late')
      return 'late'
    })
    const start = f.service.start(f.run)
    await vi.waitFor(() => expect(f.tools.openScope).toHaveBeenCalledOnce())
    f.service.release(f.run)
    opening.resolve()
    await expect(start).rejects.toThrow()
    expect(f.opened.size).toBe(0)
    f.service.dispose()
    await expect(f.service.prepare('conversation', f.run.project!.id)).rejects.toThrow()
  })

  it('drops a delayed tool result after cancellation and isolates its renderer bridge', async () => {
    const f = await fixture()
    f.service.bind(f.run)
    await f.service.start(f.run)
    const reading = deferred()
    f.tools.read.mockImplementation(async () => {
      await reading.promise
      return { status: 'ok', text: 'late', totalChars: 4, offset: 0 }
    })
    const result = f.service.tool(f.run, 'read', { path: 'material.md' })
    await vi.waitFor(() => expect(f.tools.read).toHaveBeenCalledOnce())
    f.service.release(f.run)
    reading.resolve()
    await expect(result).rejects.toThrow()
    expect(f.opened.size).toBe(0)
  })
})
