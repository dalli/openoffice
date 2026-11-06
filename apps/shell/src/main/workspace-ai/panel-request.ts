import type { WorkspacePanelRequest } from '../../shared/workspace-panel-api'
import { isWorkspaceEditQueueCommand } from '@genoffice/agent-core/workspace-editor-protocol'
import { templateOpenInputOf } from '@genoffice/electron-utils/user-skills-channels'

const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value)
const text = (value: unknown, max = 512): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= max
const id = (value: unknown) => text(value) && /^[A-Za-z0-9_-]+$/.test(value)
const positive = (value: unknown) => Number.isSafeInteger(value) && (value as number) > 0
const project = (value: unknown) => object(value) && id(value.id) && id(value.contextVersion)
const handle = (value: unknown) =>
  object(value) &&
  object(value.owner) &&
  positive(value.owner.windowId) &&
  positive(value.owner.panelContentsId) &&
  positive(value.owner.panelGeneration) &&
  positive(value.bindingId) &&
  text(value.conversationId, 2048) &&
  text(value.runId)
const origin = (value: unknown) => {
  if (!object(value)) return false
  switch (value.kind) {
    case 'workspace':
      return id(value.conversationId)
    case 'legacy-document':
    case 'project':
      return id(value.projectId) && id(value.chatId)
    case 'window-memory':
      return id(value.windowSessionId) && id(value.conversationId)
    default:
      return false
  }
}

/** Validate before any service call. Main verifies every project capability against its owner. */
export function isWorkspacePanelRequest(value: unknown): value is WorkspacePanelRequest {
  if (!object(value)) return false
  try {
    if (JSON.stringify(value).length > (value.kind === 'stream' ? 40_000_000 : 4_000_000))
      return false
  } catch {
    return false
  }
  switch (value.kind) {
    case 'open-template':
      return (
        text(value.conversationId, 2048) &&
        (value.targetTabId === undefined || text(value.targetTabId)) &&
        !!templateOpenInputOf(value.input)
      )
    case 'entry-ack':
      return (
        text(value.entryId, 256) &&
        typeof value.accepted === 'boolean' &&
        (value.conversationId === undefined || text(value.conversationId, 2048))
      )
    case 'entries-ready':
      return typeof value.ready === 'boolean'
    case 'stream-cancel':
      return handle(value.run) && id(value.requestId)
    case 'edit-queue':
    case 'template': {
      const d = value.document
      return (
        object(d) &&
        positive(d.windowId) &&
        text(d.tabId) &&
        text(d.documentId) &&
        positive(d.generation) &&
        d.kind === 'docs' &&
        (value.kind === 'edit-queue'
          ? isWorkspaceEditQueueCommand(value.command)
          : ['info', 'preview', 'as-is'].includes(value.mode as string))
      )
    }
    case 'stream': {
      const r = value.request
      return (
        handle(value.run) &&
        object(r) &&
        (r.purpose === undefined ||
          (r.purpose === 'writer' && Array.isArray(r.tools) && r.tools.length === 0)) &&
        id(r.requestId) &&
        id(r.sessionId) &&
        typeof r.system === 'string' &&
        Array.isArray(r.messages) &&
        r.messages.every(
          (m) =>
            object(m) &&
            (m.role === 'tool'
              ? Array.isArray(m.results)
              : (m.role === 'user' || m.role === 'assistant') && typeof m.text === 'string'),
        ) &&
        Array.isArray(r.tools) &&
        r.tools.every(
          (t) =>
            object(t) &&
            typeof t.name === 'string' &&
            /^[A-Za-z0-9_-]{1,64}$/.test(t.name) &&
            typeof t.description === 'string' &&
            object(t.inputSchema),
        )
      )
    }
    case 'drive-status':
      return true
    case 'drive-pick':
      return text(value.conversationId, 2048)
    case 'drive-check':
      return handle(value.run)
    case 'drive-search':
      return (
        handle(value.run) &&
        text(value.query, 4000) &&
        Number.isSafeInteger(value.limit) &&
        (value.limit as number) >= 1 &&
        (value.limit as number) <= 20
      )
    case 'drive-read':
      return (
        handle(value.run) &&
        typeof value.fileId === 'string' &&
        /^[A-Za-z0-9_-]{1,200}$/.test(value.fileId) &&
        Number.isSafeInteger(value.offset) &&
        (value.offset as number) >= 0
      )
    case 'attachment-paste':
      return (
        text(value.conversationId, 2048) &&
        value.bytes instanceof ArrayBuffer &&
        value.bytes.byteLength > 0 &&
        value.bytes.byteLength <= 5 * 1024 * 1024 &&
        ['png', 'jpg', 'jpeg', 'gif', 'webp'].includes(value.ext as string)
      )
    case 'attachment-pick':
      return text(value.conversationId, 2048)
    case 'attachment-add':
      return (
        text(value.conversationId, 2048) &&
        Array.isArray(value.paths) &&
        value.paths.length > 0 &&
        value.paths.length <= 20 &&
        value.paths.every((p) => text(p, 32768) && !p.includes('\0'))
      )
    case 'attachment-list':
      return handle(value.run)
    case 'attachment-read':
    case 'attachment-image':
      return (
        handle(value.run) &&
        Number.isSafeInteger(value.index) &&
        (value.index as number) >= 0 &&
        (value.index as number) < 20 &&
        (value.kind === 'attachment-image' ||
          (Number.isSafeInteger(value.offset) && (value.offset as number) >= 0))
      )
    case 'catalog':
    case 'startNext':
    case 'disconnect':
      return true
    case 'create':
      return value.projectId === undefined || id(value.projectId)
    case 'project-prepare':
      return text(value.conversationId, 2048) && id(value.projectId)
    case 'project-discard':
      return text(value.conversationId, 2048) && project(value.binding)
    case 'project-context':
      return handle(value.run)
    case 'project-tool': {
      if (!handle(value.run) || !object(value.args)) return false
      const a = value.args
      // Scope comes exclusively from main. Callers may only supply existing tool arguments.
      const keys: Record<string, string[]> = {
        listFiles: ['prefix', 'offset', 'limit'],
        search: ['query', 'exts', 'offset', 'limit'],
        read: ['relPath', 'offset'],
        remember: ['change'],
      }
      const allowed = keys[value.method as string]
      if (!Array.isArray(allowed) || Object.keys(a).some((key) => !allowed.includes(key)))
        return false
      if (
        [a.offset, a.limit].some(
          (n) => n !== undefined && (!Number.isSafeInteger(n) || (n as number) < 0),
        )
      )
        return false
      if (value.method === 'read') return text(a.relPath, 4096)
      if (value.method === 'search')
        return (
          text(a.query, 4000) &&
          (a.exts === undefined ||
            (Array.isArray(a.exts) && a.exts.length <= 64 && a.exts.every((e) => text(e, 32))))
        )
      if (value.method === 'listFiles')
        return a.prefix === undefined || (typeof a.prefix === 'string' && a.prefix.length <= 4096)
      const c = a.change
      return (
        object(c) &&
        ((['add', 'remove'].includes(c.op as string) && text(c.text, 8000)) ||
          (c.op === 'replace' && text(c.old, 8000) && text(c.new, 8000)))
      )
    }
    case 'layout':
      return (
        value.panel === null ||
        (object(value.panel) &&
          ['left', 'right'].includes(value.panel.side as string) &&
          typeof value.panel.open === 'boolean' &&
          typeof value.panel.width === 'number' &&
          Number.isFinite(value.panel.width))
      )
    case 'capture':
      return text(value.tabId)
    case 'refresh':
      return text(value.captureId)
    case 'reference':
      return text(value.captureId) && ['document', 'selection'].includes(value.scope as string)
    case 'state':
      return text(value.runId)
    case 'load':
      return origin(value.origin)
    case 'append': {
      const m = value.message
      return (
        origin(value.origin) &&
        object(m) &&
        ['user', 'assistant'].includes(m.role as string) &&
        typeof m.text === 'string' &&
        m.text.length <= 200_000 &&
        (m.tools === undefined ||
          (Array.isArray(m.tools) &&
            m.tools.length <= 64 &&
            m.tools.every(
              (t) =>
                object(t) &&
                text(t.name, 200) &&
                typeof t.summary === 'string' &&
                [t.input, t.output].every((s) => s === undefined || typeof s === 'string') &&
                (t.isError === undefined || typeof t.isError === 'boolean'),
            ))) &&
        (m.attachments === undefined ||
          (Array.isArray(m.attachments) &&
            m.attachments.length <= 32 &&
            m.attachments.every(
              (a) =>
                object(a) &&
                typeof a.name === 'string' &&
                [a.path, a.ext].every((s) => s === undefined || typeof s === 'string') &&
                (a.sizeBytes === undefined ||
                  (typeof a.sizeBytes === 'number' && Number.isFinite(a.sizeBytes))),
            ))) &&
        (m.scope === undefined ||
          (object(m.scope) &&
            typeof m.scope.label === 'string' &&
            (m.scope.text === undefined || typeof m.scope.text === 'string'))) &&
        (m.skills === undefined || (Array.isArray(m.skills) && m.skills.length <= 32)) &&
        (m.fileRef === undefined || typeof m.fileRef === 'string')
      )
    }
    case 'enqueue': {
      const r = value.request
      return (
        object(r) &&
        text(r.conversationId, 2048) &&
        text(r.runId) &&
        text(r.targetId) &&
        Array.isArray(r.referenceIds) &&
        r.referenceIds.length <= 20 &&
        r.referenceIds.every((i) => text(i)) &&
        (r.project === null || project(r.project)) &&
        (value.attachments === undefined ||
          (Array.isArray(value.attachments) &&
            value.attachments.length <= 20 &&
            value.attachments.every(
              (a) =>
                object(a) &&
                text(a.path, 32768) &&
                text(a.name, 4096) &&
                text(a.ext, 32) &&
                Number.isSafeInteger(a.sizeBytes) &&
                (a.sizeBytes as number) >= 0,
            )))
      )
    }
    case 'complete':
    case 'cancel':
    case 'rollbackForRun':
      return Array.isArray(value.args) && value.args.length === 1 && handle(value.args[0])
    case 'rollback':
      return (
        Array.isArray(value.args) &&
        value.args.length === 1 &&
        handle(value.args[0]) &&
        text(value.args[0].rollbackId)
      )
    case 'writer-update': {
      if (
        !Array.isArray(value.args) ||
        value.args.length !== 3 ||
        !handle(value.args[0]) ||
        !text(value.args[1])
      )
        return false
      const update = value.args[2]
      if (!object(update)) return false
      if (update.kind === 'ready') return true
      if (update.kind === 'progress')
        return typeof update.text === 'string' && update.text.length <= 210_000
      return (
        update.kind === 'result' &&
        typeof update.ok === 'boolean' &&
        (update.text === undefined ||
          (typeof update.text === 'string' && update.text.length <= 210_000)) &&
        (update.truncated === undefined || typeof update.truncated === 'boolean') &&
        (update.error === undefined ||
          (typeof update.error === 'string' && update.error.length <= 20_000))
      )
    }
    case 'execute': {
      if (!Array.isArray(value.args) || value.args.length !== 3 || !handle(value.args[0]))
        return false
      const [, d, c] = value.args
      return (
        object(d) &&
        positive(d.windowId) &&
        text(d.tabId) &&
        text(d.documentId) &&
        positive(d.generation) &&
        ['docs', 'markdown', 'hwp'].includes(d.kind as string) &&
        object(c) &&
        text(c.id) &&
        text(c.name, 200) &&
        object(c.input)
      )
    }
    case 'verifyResponse':
      return (
        Array.isArray(value.args) &&
        value.args.length === 3 &&
        handle(value.args[0]) &&
        typeof value.args[1] === 'string' &&
        value.args[1].length <= 200_000 &&
        Array.isArray(value.args[2]) &&
        value.args[2].length <= 512 &&
        value.args[2].every((c) => object(c) && text(c.name, 200) && typeof c.ok === 'boolean')
      )
    default:
      return false
  }
}
