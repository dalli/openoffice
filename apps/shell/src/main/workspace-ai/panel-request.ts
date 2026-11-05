import type { WorkspacePanelRequest } from '../../shared/workspace-panel-api'

const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value)
const text = (value: unknown, max = 512): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= max
const id = (value: unknown) => text(value) && /^[A-Za-z0-9_-]+$/.test(value)
const positive = (value: unknown) => Number.isSafeInteger(value) && (value as number) > 0
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

/** Validate before any service call. Project bindings await a main-owned preparation capability. */
export function isWorkspacePanelRequest(value: unknown): value is WorkspacePanelRequest {
  if (!object(value)) return false
  try {
    if (JSON.stringify(value).length > 4_000_000) return false
  } catch {
    return false
  }
  switch (value.kind) {
    case 'catalog':
    case 'create':
    case 'startNext':
    case 'disconnect':
      return true
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
        r.project === null
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
