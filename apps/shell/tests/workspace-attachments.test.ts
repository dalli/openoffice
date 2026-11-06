import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { setOpenFilesSource } from '@genoffice/electron-utils/unsaved-files'
import { parseFileToText } from '@genoffice/file-parse'
import {
  WorkspaceAttachments,
  type WorkspaceAttachmentServices,
} from '../src/main/workspace-ai/attachments'
import type {
  WorkspaceRunHandle,
  WorkspaceStartedRun,
} from '../src/main/workspace-ai/execution-controller'
import type { AttachmentMeta } from '../../docs/src/shared/ipc'
import { prepareWorkspaceAttachments } from '../src/renderer/src/workspace-ai/attachment-preparation'
import type { WorkspaceComposer } from '../src/renderer/src/workspace-ai/conversations'

const cleanup: Array<() => void | Promise<void>> = []
afterEach(async () => {
  setOpenFilesSource(null)
  for (const off of cleanup.splice(0).reverse()) await off()
})
const handle = (bindingId = 1): WorkspaceRunHandle => ({
  bindingId,
  conversationId: 'conversation',
  runId: `run-${bindingId}`,
  owner: { windowId: 1, panelContentsId: 10, panelGeneration: 1 },
})
async function fixture(ext = 'md') {
  const root = await mkdtemp(join(tmpdir(), 'workspace-attachments-'))
  cleanup.push(() => rm(root, { recursive: true, force: true }))
  const path = join(root, `source.${ext}`)
  const text =
    'Visible <span class="redact-span" data-redaction="" data-label="private">HIDDEN</span> material'
  await writeFile(path, text)
  const file: AttachmentMeta = {
    path,
    name: basename(path),
    ext,
    sizeBytes: (await stat(path)).size,
  }
  const services: WorkspaceAttachmentServices = {
    add: vi.fn(() => ({ accepted: [file], rejected: [] })),
    pick: vi.fn(async () => ({ accepted: [file], rejected: [] })),
    read: vi.fn(async (path, offset, maxChars) => {
      const parsed = await parseFileToText(path)
      return parsed.ok && parsed.kind === 'text'
        ? {
            ok: true,
            offset,
            totalChars: parsed.text!.length,
            text: parsed.text!.slice(offset, offset + maxChars),
          }
        : { ok: false, error: 'Unreadable' }
    }),
    image: vi.fn(() => ({ ok: true, base64: 'cGl4ZWw=', mime: 'image/png' })),
  }
  const attachments = new WorkspaceAttachments(services)
  cleanup.push(() => attachments.dispose())
  return { attachments, services, file, path, text }
}

describe('workspace attachment capabilities', () => {
  it('reads only explicit run attachments through the existing redacting parser and leaves files intact', async () => {
    const f = await fixture()
    await f.attachments.add('conversation', [f.path])
    await f.attachments.bind(handle(), [f.file])
    expect(await f.attachments.list(handle())).toEqual([f.file])
    const result = await f.attachments.read(handle(), 0, 0)
    expect(result.text).toContain('{{private}}')
    expect(result.text).not.toContain('HIDDEN')
    expect(await readFile(f.path, 'utf8')).toBe(f.text)
    expect(f.services.read).toHaveBeenCalledWith(f.path, 0, 24_000)
  })

  it('rejects paths never selected, forged metadata, another conversation and oversized requests', async () => {
    const f = await fixture()
    await expect(f.attachments.bind(handle(), [f.file])).rejects.toThrow()
    await f.attachments.add('conversation', [f.path])
    for (const changed of [
      { name: 'forged' },
      { ext: 'png' },
      { sizeBytes: 0 },
      { path: '/other' },
    ])
      await expect(f.attachments.bind(handle(), [{ ...f.file, ...changed }])).rejects.toThrow()
    await expect(
      f.attachments.bind({ ...handle(), conversationId: 'other' }, [f.file]),
    ).rejects.toThrow()
    await expect(f.attachments.bind(handle(), Array(21).fill(f.file))).rejects.toThrow()
    await expect(f.attachments.bind(handle(), [f.file, f.file])).rejects.toThrow()
    expect(f.services.read).not.toHaveBeenCalled()
  })

  it('rejects changed owners, generations, run ids, conversations and indices', async () => {
    const f = await fixture()
    await f.attachments.add('conversation')
    await f.attachments.bind(handle(), [f.file])
    for (const run of [
      { ...handle(), owner: { ...handle().owner, windowId: 2 } },
      { ...handle(), owner: { ...handle().owner, panelGeneration: 2 } },
      { ...handle(), owner: { ...handle().owner, panelContentsId: 20 } },
      { ...handle(), runId: 'other' },
      { ...handle(), conversationId: 'other' },
      handle(2),
    ])
      await expect(f.attachments.read(run, 0, 0)).rejects.toThrow()
    await expect(f.attachments.read(handle(), 1, 0)).rejects.toThrow()
    await expect(f.attachments.image(handle(), 0)).rejects.toThrow()
    expect(f.services.read).not.toHaveBeenCalled()
    expect(f.services.image).not.toHaveBeenCalled()
  })

  it('checks unsaved changes before binding, before reading, and after an asynchronous read', async () => {
    const f = await fixture()
    await f.attachments.add('conversation')
    let dirty = true
    setOpenFilesSource(() => [{ filePath: f.path, dirty }])
    await expect(f.attachments.bind(handle(), [f.file])).rejects.toThrow('unsaved')
    dirty = false
    await f.attachments.bind(handle(), [f.file])
    dirty = true
    await expect(f.attachments.read(handle(), 0, 0)).rejects.toThrow('unsaved')
    expect(f.services.read).not.toHaveBeenCalled()
    dirty = false
    vi.mocked(f.services.read).mockImplementation(async () => {
      dirty = true
      return { ok: true, text: 'STALE' }
    })
    await expect(f.attachments.read(handle(), 0, 0)).rejects.toThrow('unsaved')
  })

  it('rejects disk changes while queued and during parsing', async () => {
    const f = await fixture()
    await f.attachments.add('conversation')
    await f.attachments.bind(handle(), [f.file])
    await writeFile(f.path, 'changed')
    await expect(f.attachments.list(handle())).rejects.toThrow()
    expect(f.services.read).not.toHaveBeenCalled()
    await f.attachments.bind(handle(2), [f.file])
    vi.mocked(f.services.read).mockImplementation(async () => {
      await writeFile(f.path, 'changed during parsing')
      return { ok: true, text: 'STALE' }
    })
    await expect(f.attachments.read(handle(2), 0, 0)).rejects.toThrow()
  })

  it('discards reads and picker replies that finish after release or panel closure', async () => {
    const f = await fixture()
    await f.attachments.add('conversation')
    await f.attachments.bind(handle(), [f.file])
    let finish!: () => void
    vi.mocked(f.services.read).mockImplementation(
      () =>
        new Promise((resolve) => {
          finish = () => resolve({ ok: true, text: 'LATE' })
        }),
    )
    const pending = f.attachments.read(handle(), 0, 0)
    await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
    f.attachments.release(handle())
    finish()
    await expect(pending).rejects.toThrow()
    let pick!: () => void
    vi.mocked(f.services.pick).mockImplementation(
      () =>
        new Promise((resolve) => {
          pick = () => resolve({ accepted: [f.file], rejected: [] })
        }),
    )
    const choosing = f.attachments.add('conversation')
    f.attachments.dispose()
    pick()
    await expect(choosing).rejects.toThrow()
  })

  it('accepts clipboard images through the same grant and run checks, rejecting oversized or non-image data before writing', async () => {
    const f = await fixture('png')
    f.services.paste = vi.fn(() => ({ accepted: [f.file], rejected: [] }))
    const bytes = new Uint8Array([1, 2, 3]).buffer
    expect(f.attachments.paste('conversation', bytes, 'png')?.accepted).toEqual([f.file])
    await f.attachments.bind(handle(), [f.file])
    expect((await f.attachments.image(handle(), 0)).base64).toBe('cGl4ZWw=')
    for (const [data, ext] of [
      [new ArrayBuffer(0), 'png'],
      [new ArrayBuffer(5 * 1024 * 1024 + 1), 'png'],
      [bytes, 'html'],
    ] as const)
      expect(() => f.attachments.paste('conversation', data, ext)).toThrow()
    expect(f.services.paste).toHaveBeenCalledTimes(1)
    f.attachments.dispose()
    expect(() => f.attachments.paste('conversation', bytes, 'png')).toThrow()
  })

  it('never reads a draw.io PNG as a raw image', async () => {
    const f = await fixture('drawio.png')
    await f.attachments.add('conversation')
    await f.attachments.bind(handle(), [f.file])
    await expect(f.attachments.image(handle(), 0)).rejects.toThrow()
    expect(f.services.image).not.toHaveBeenCalled()
  })
})

describe('workspace attachment preparation', () => {
  const file: AttachmentMeta = {
    path: '/material.md',
    name: 'material.md',
    ext: 'md',
    sizeBytes: 20,
  }
  const picture: AttachmentMeta = {
    path: '/picture.png',
    name: 'picture.png',
    ext: 'png',
    sizeBytes: 10,
  }
  const input = (attachments = [file, picture]) => ({ attachments }) as WorkspaceComposer
  const run = (signal = new AbortController().signal) =>
    ({ binding: { ...handle(), project: null }, signal }) as WorkspaceStartedRun
  const client = () => ({
    attachmentList: vi.fn(async () => [file, picture]),
    attachmentRead: vi.fn(async () => ({ ok: true, offset: 0, totalChars: 5, text: 'hello' })),
    attachmentImage: vi.fn(async () => ({ ok: true, base64: 'cGl4ZWw=', mime: 'image/png' })),
  })
  it('reuses file paging and localized summaries, loads images once, and retires the skill on disposal', async () => {
    const c = client()
    const prepared = await prepareWorkspaceAttachments(c, input(), run(), () => 'en')
    expect(prepared.images).toEqual([{ base64: 'cGl4ZWw=', mime: 'image/png' }])
    const skill = prepared.skills[0]!
    expect(skill.buildContext?.()).toContain('material.md')
    const call = { id: 'read', name: 'read_attachment', input: { index: 0, offset: 2.8 } }
    expect((await skill.executeTool(call)).summary).toBe('Read attachment material.md')
    expect(c.attachmentRead).toHaveBeenCalledWith(run().binding, 0, 2)
    await skill.executeTool({ ...call, input: { index: 1 } })
    expect(c.attachmentRead).toHaveBeenCalledTimes(1)
    expect(c.attachmentImage).toHaveBeenCalledTimes(1)
    prepared.dispose()
    await expect(skill.executeTool(call)).rejects.toThrow()
  })
  it('does no reads for an empty composer and rejects mismatched metadata and unreadable images', async () => {
    const c = client()
    await prepareWorkspaceAttachments(c, input([]), run(), () => 'en')
    expect(c.attachmentList).not.toHaveBeenCalled()
    await expect(
      prepareWorkspaceAttachments(c, input([{ ...file, path: '/forged' }]), run(), () => 'en'),
    ).rejects.toThrow()
    c.attachmentImage.mockResolvedValue({ ok: false, base64: '', mime: '' })
    await expect(prepareWorkspaceAttachments(c, input(), run(), () => 'en')).rejects.toThrow()
  })
  it('drops an image reply arriving after cancellation', async () => {
    const c = client()
    const abort = new AbortController()
    c.attachmentImage.mockImplementation(async () => {
      abort.abort()
      return { ok: true, base64: 'LATE', mime: 'image/png' }
    })
    await expect(
      prepareWorkspaceAttachments(c, input(), run(abort.signal), () => 'en'),
    ).rejects.toThrow()
  })
})
