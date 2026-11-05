import { randomUUID } from 'node:crypto'
import type { BrowserWindow, IpcMain, IpcMainInvokeEvent } from 'electron'
import type { ChatMessage } from '@genoffice/project-store'
import {
  WORKSPACE_MEMORY_CHANNELS as C,
  type WorkspaceMemoryRequest,
} from '../../shared/workspace-memory-api'
import type { WorkspaceConversationHistory } from './conversation-history'
import { isWorkspacePanelRequest } from './panel-request'

interface Services {
  window: BrowserWindow
  history: WorkspaceConversationHistory
}
interface Host {
  id: string
  services: Services
  frame: IpcMainInvokeEvent['senderFrame']
  pending: Map<
    string,
    {
      resolve(value: readonly ChatMessage[]): void
      reject(): void
      timer: ReturnType<typeof setTimeout>
      kind: WorkspaceMemoryRequest['kind']
    }
  >
  off(): void
}

/** Reads and writes at the original home host. Main never caches a transcript or creates a file. */
export function registerWorkspaceMemoryIpc(
  ipc: Pick<IpcMain, 'handle' | 'removeHandler'>,
  servicesFor: (event: IpcMainInvokeEvent) => Services | null,
  changed: () => void,
  timeoutMs = 5000,
): { dispose(): void } {
  let host: Host | null = null
  const owns = (h: Host, event: IpcMainInvokeEvent) =>
    host === h &&
    !h.services.window.isDestroyed() &&
    !event.sender.isDestroyed() &&
    event.sender === h.services.window.webContents &&
    event.senderFrame === event.sender.mainFrame &&
    event.senderFrame === h.frame
  const sendRelease = (h: Host) => {
    if (h.services.window.isDestroyed()) return
    const wc = h.services.window.webContents
    if (!wc.isDestroyed() && wc.mainFrame === h.frame)
      wc.send(C.request, { hostId: h.id, id: randomUUID(), kind: 'release' })
  }
  const close = (h: Host) => {
    if (host !== h) return
    sendRelease(h)
    host = null
    h.off()
    for (const p of h.pending.values()) {
      clearTimeout(p.timer)
      p.reject()
    }
    h.pending.clear()
    changed()
  }
  const request = (
    h: Host,
    operation: Omit<WorkspaceMemoryRequest, 'id' | 'hostId'> & { message?: unknown },
  ): Promise<readonly ChatMessage[]> => {
    if (host !== h || h.pending.size >= 16)
      return Promise.reject(new Error('Memory host unavailable'))
    const id = randomUUID()
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => close(h), timeoutMs)
      h.pending.set(id, {
        resolve,
        reject: () => reject(new Error('Memory host unavailable')),
        timer,
        kind: operation.kind,
      })
      try {
        h.services.window.webContents.send(C.request, { ...operation, id, hostId: h.id })
      } catch {
        close(h)
      }
    })
  }
  ipc.handle(C.register, (event) => {
    const services = servicesFor(event)
    if (
      !services ||
      services.window.isDestroyed() ||
      event.sender.isDestroyed() ||
      event.sender !== services.window.webContents ||
      event.senderFrame !== event.sender.mainFrame
    )
      return null
    if (host) close(host)
    const id = randomUUID()
    const closed = () => close(h)
    const navigate = (_event: unknown, _url: string, inPlace: boolean, mainFrame: boolean) => {
      if (mainFrame && !inPlace) closed()
    }
    const detach = services.history.attachMemory(id, {
      load: () => request(h, { kind: 'load' }),
      append: async (message) => {
        await request(h, { kind: 'append', message })
      },
      release: () => sendRelease(h),
    })
    const h: Host = {
      id,
      services,
      frame: event.senderFrame,
      pending: new Map(),
      off() {
        detach()
        event.sender.removeListener('destroyed', closed)
        event.sender.removeListener('render-process-gone', closed)
        event.sender.removeListener('did-start-navigation', navigate)
        services.window.removeListener('closed', closed)
      },
    }
    host = h
    event.sender.on('destroyed', closed)
    event.sender.on('render-process-gone', closed)
    event.sender.on('did-start-navigation', navigate)
    services.window.on('closed', closed)
    changed()
    return id
  })
  ipc.handle(C.unregister, (event, id: unknown) => {
    if (host && id === host.id && owns(host, event)) close(host)
  })
  ipc.handle(C.reply, (event, hostId: unknown, id: unknown, raw: unknown) => {
    const h = host
    if (!h || hostId !== h.id || !owns(h, event) || typeof id !== 'string') return
    const p = h.pending.get(id)
    if (!p) return
    h.pending.delete(id)
    clearTimeout(p.timer)
    const reply = raw as { ok?: boolean; messages?: ChatMessage[] } | null
    const messages = reply?.messages
    let valid = reply?.ok === true
    if (p.kind === 'load') {
      valid =
        valid &&
        Array.isArray(messages) &&
        messages.length <= 200 &&
        messages.every(
          (m, index) =>
            m &&
            Number.isSafeInteger(m.seq) &&
            m.seq >= 0 &&
            (index === 0 || m.seq > messages[index - 1]!.seq) &&
            typeof m.ts === 'string' &&
            Number.isFinite(Date.parse(m.ts)) &&
            isWorkspacePanelRequest({
              kind: 'append',
              origin: { kind: 'workspace', conversationId: 'validation' },
              message: m,
            }),
        )
      try {
        valid = valid && JSON.stringify(messages).length <= 4_000_000
      } catch {
        valid = false
      }
    }
    if (valid) p.resolve(messages ?? [])
    else {
      if (p.kind === 'load') sendRelease(h)
      p.reject()
    }
  })
  return {
    dispose() {
      if (host) close(host)
      for (const channel of [C.register, C.unregister, C.reply]) ipc.removeHandler(channel)
    },
  }
}
