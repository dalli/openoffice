import { randomUUID } from 'node:crypto'
import type { BrowserWindow, IpcMain, IpcMainInvokeEvent } from 'electron'
import type { ChatSender } from '../projects/chat-store'
import type { TabManager } from '../tab-manager'
import { workspaceConversationKey } from '../../shared/workspace-ai-api'
import {
  WORKSPACE_PANEL_CHANNELS as C,
  type WorkspacePanelApi,
  type WorkspacePanelCatalog,
  type WorkspacePanelEvent,
  type WorkspacePanelReply,
  type WorkspacePanelRequest,
  type WorkspacePanelSession,
  type WorkspacePanelSource,
  type WorkspacePanelValues,
} from '../../shared/workspace-panel-api'
import { WorkspaceExecutionError, type WorkspaceExecutionController } from './execution-controller'
import type { WorkspaceConversationHistory } from './conversation-history'
import { isWorkspacePanelRequest } from './panel-request'

export interface WorkspacePanelServices {
  window: BrowserWindow
  tabs: TabManager
  history: WorkspaceConversationHistory
  sources(): Promise<WorkspacePanelSource[]>
  /** Main derives the saved path and editor from the bound target, never from a message. */
  sender(tabId?: string): ChatSender
}

interface Session {
  info: WorkspacePanelSession
  services: WorkspacePanelServices
  execution: WorkspaceExecutionController
  frame: IpcMainInvokeEvent['senderFrame']
  opened: Set<string>
  origins: Map<string, WorkspacePanelSource['origin']>
  targets: Map<string, string>
  rebindings: Map<string, WorkspacePanelSource['origin']>
  writing: Set<string>
  off(): void
}

/** Private shell bridge. No model, provider, filesystem path or arbitrary channel operations. */
export function registerWorkspacePanelIpc(
  ipc: Pick<IpcMain, 'handle' | 'removeHandler'>,
  servicesFor: (event: IpcMainInvokeEvent) => WorkspacePanelServices | null,
): {
  changed(): void
  historyChanged(origin: WorkspacePanelSource['origin']): void
  rebound(
    previous: WorkspacePanelSource['origin'],
    origin: Extract<WorkspacePanelSource['origin'], { kind: 'legacy-document' }>,
  ): void
  dispose(): void
} {
  let session: Session | null = null
  let generation = 0
  const fail = (reason: 'wrong-owner' | 'invalid-request' | 'panel-closed'): never => {
    throw new WorkspaceExecutionError(reason)
  }
  const emit = (s: Session, event: WorkspacePanelEvent) => {
    if (session !== s || s.services.window.isDestroyed()) return
    const contents = s.services.window.webContents
    if (!contents.isDestroyed()) contents.send(C.event, event)
  }
  const close = (s: Session) => {
    if (session !== s) return
    emit(s, { sessionId: s.info.id, kind: 'closed' })
    session = null
    s.off()
    s.services.history.releaseMemory()
    s.execution.dispose()
    if (!s.services.window.isDestroyed()) s.services.tabs.setWorkspacePanel(null)
  }
  const owns = (s: Session, event: IpcMainInvokeEvent) =>
    !s.services.window.isDestroyed() &&
    event.sender === s.services.window.webContents &&
    !event.sender.isDestroyed() &&
    event.senderFrame === event.sender.mainFrame &&
    event.senderFrame === s.frame
  const live = (s: Session) => {
    if (session !== s) fail('panel-closed')
  }
  const guarded = async <T>(operation: () => T | Promise<T>): Promise<WorkspacePanelReply<T>> => {
    try {
      return { ok: true, value: await operation() }
    } catch (error) {
      // Do not leak disk paths, provider errors or project file contents through Electron errors.
      return {
        ok: false,
        reason: error instanceof WorkspaceExecutionError ? error.kind : 'service-error',
      }
    }
  }
  const resolveOrigin = (s: Session, origin: WorkspacePanelSource['origin']) => {
    const visited = new Set<string>()
    while (s.rebindings.has(workspaceConversationKey(origin))) {
      const key = workspaceConversationKey(origin)
      if (visited.has(key)) fail('invalid-request')
      visited.add(key)
      origin = s.rebindings.get(key)!
    }
    return origin
  }
  const source = async (
    s: Session,
    origin: Parameters<WorkspaceConversationHistory['load']>[0],
  ) => {
    const requestedKey = workspaceConversationKey(origin)
    const available = await s.services.sources()
    live(s)
    origin = resolveOrigin(s, origin)
    const key = workspaceConversationKey(origin)
    if (!available.some((item) => workspaceConversationKey(item.origin) === key))
      fail('invalid-request')
    return { key, requestedKey, origin }
  }
  const catalog = async (s: Session): Promise<WorkspacePanelCatalog> => {
    const sources = await s.services.sources()
    live(s)
    const split = s.services.tabs.splitState()
    const documents = s.services.tabs.workspaceDocuments().map((tab) => ({
      id: tab.id,
      title: tab.title,
      kind: tab.kind,
      active: tab.active,
      available: tab.available,
      ...(split?.aTabId === tab.id
        ? { pane: split.orientation === 'side' ? ('left' as const) : ('top' as const) }
        : split?.bTabId === tab.id
          ? { pane: split.orientation === 'side' ? ('right' as const) : ('bottom' as const) }
          : {}),
    }))
    return { documents, sources }
  }
  ipc.handle(C.connect, (event): ReturnType<WorkspacePanelApi['connect']> =>
    guarded(() => {
      if (event.sender.isDestroyed() || event.senderFrame !== event.sender.mainFrame)
        fail('wrong-owner')
      const services = servicesFor(event) ?? fail('wrong-owner')
      if (
        event.sender !== services.window.webContents ||
        event.senderFrame !== event.sender.mainFrame ||
        event.sender.isDestroyed()
      )
        fail('wrong-owner')
      if (session) {
        if (!owns(session, event)) fail('wrong-owner')
        close(session)
      }
      const info: WorkspacePanelSession = {
        id: randomUUID(),
        owner: {
          windowId: services.window.id,
          panelContentsId: event.sender.id,
          panelGeneration: ++generation,
        },
      }
      const execution = services.tabs.workspaceAiForPanel(info.owner.panelGeneration, {
        onStopped: (run) => {
          if (session?.info === info) emit(session, { sessionId: info.id, kind: 'stopped', run })
        },
        onDisposed: () => {
          if (session?.info === info) close(session)
        },
      })
      const closed = () => {
        if (session?.info === info) close(session)
      }
      const navigation = (_event: unknown, _url: string, inPlace: boolean, mainFrame: boolean) => {
        if (mainFrame && !inPlace) closed()
      }
      const offLayout = services.tabs.onWorkspaceLayoutChanged((layout) => {
        if (session?.info === info) emit(session, { sessionId: info.id, kind: 'layout', layout })
      })
      event.sender.on('did-start-navigation', navigation)
      event.sender.on('render-process-gone', closed)
      event.sender.on('destroyed', closed)
      services.window.on('closed', closed)
      session = {
        info,
        services,
        execution,
        frame: event.senderFrame,
        opened: new Set(),
        origins: new Map(),
        targets: new Map(),
        rebindings: new Map(),
        writing: new Set(),
        off: () => {
          offLayout()
          event.sender.removeListener('did-start-navigation', navigation)
          event.sender.removeListener('render-process-gone', closed)
          event.sender.removeListener('destroyed', closed)
          services.window.removeListener('closed', closed)
        },
      }
      return info
    }),
  )

  const dispatch = async (
    s: Session,
    r: WorkspacePanelRequest,
  ): Promise<WorkspacePanelValues[keyof WorkspacePanelValues]> => {
    const e = s.execution
    switch (r.kind) {
      case 'catalog':
        return catalog(s)
      case 'layout':
        return s.services.tabs.setWorkspacePanel(r.panel)
      case 'capture':
        return e.capture(r.tabId)
      case 'reference': {
        const { id, binding } = await e.readReference(r.captureId, { scope: r.scope })
        return { id, binding }
      }
      case 'create':
        return { origin: s.services.history.createLocal(), title: '' }
      case 'load': {
        const ref = await source(s, r.origin)
        const history = await s.services.history.load(ref.origin)
        live(s)
        const origin = resolveOrigin(s, ref.origin)
        const key = workspaceConversationKey(origin)
        const requestedKey = ref.requestedKey
        s.opened.add(key)
        s.opened.add(requestedKey)
        s.origins.set(key, origin)
        s.origins.set(requestedKey, origin)
        return { ...history, origin }
      }
      case 'append': {
        const { key, requestedKey, origin } = await source(s, r.origin)
        if (!s.opened.has(key)) fail('invalid-request')
        s.writing.add(key)
        try {
          return await s.services.history.append(
            origin,
            r.message,
            s.services.sender(s.targets.get(key) ?? s.targets.get(requestedKey)),
          )
        } finally {
          s.writing.delete(key)
        }
      }
      case 'enqueue': {
        if (!s.opened.has(r.request.conversationId)) fail('invalid-request')
        const origin = s.origins.get(r.request.conversationId) ?? fail('invalid-request')
        await source(s, origin)
        const binding = e.enqueue(r.request)
        return binding
      }
      case 'startNext': {
        const run = await e.startNext()
        if (run) {
          try {
            const origin = s.origins.get(run.binding.conversationId) ?? fail('invalid-request')
            const { key } = await source(s, origin)
            s.targets.set(key, run.binding.target.document.tabId)
            s.targets.set(run.binding.conversationId, run.binding.target.document.tabId)
          } catch (error) {
            await e.cancel(run.binding)
            throw error
          }
        }
        return run
          ? {
              binding: run.binding,
              target: run.target,
              references: run.references,
              aborted: run.signal.aborted,
            }
          : null
      }
      case 'execute':
        return e.execute(...r.args)
      case 'verifyResponse':
        return e.verifyResponse(...r.args)
      case 'complete':
        return e.complete(...r.args)
      case 'cancel':
        return e.cancel(...r.args)
      case 'state':
        return e.state(r.runId)
      case 'rollbackForRun':
        return e.rollbackForRun(...r.args)
      case 'rollback':
        return e.rollback(...r.args)
      case 'disconnect':
        close(s)
        return
    }
  }
  ipc.handle(C.request, (event, sessionId: unknown, raw: unknown) =>
    guarded(async () => {
      const s = session ?? fail('panel-closed')
      if (sessionId !== s.info.id || !owns(s, event)) fail('wrong-owner')
      const request = isWorkspacePanelRequest(raw) ? raw : fail('invalid-request')
      const value = await dispatch(s, request)
      if (request.kind !== 'disconnect') live(s)
      return value
    }),
  )
  return {
    historyChanged: (origin) => {
      if (session && !session.writing.has(workspaceConversationKey(origin)))
        emit(session, { sessionId: session.info.id, kind: 'history-changed', origin })
    },
    rebound: (previous, origin) => {
      const s = session
      if (!s) return
      const before = workspaceConversationKey(previous)
      const after = workspaceConversationKey(origin)
      if (before === after) return
      for (const [key, value] of s.rebindings)
        if (workspaceConversationKey(value) === before) s.rebindings.set(key, origin)
      s.rebindings.delete(after)
      s.rebindings.set(before, origin)
      if (s.opened.has(before)) {
        s.opened.add(after)
        s.origins.set(before, origin)
        s.origins.set(after, origin)
        const target = s.targets.get(before)
        if (target) s.targets.set(after, target)
      }
      emit(s, { sessionId: s.info.id, kind: 'history-rebound', previous, origin })
    },
    changed: () => {
      if (session) emit(session, { sessionId: session.info.id, kind: 'catalog-changed' })
    },
    dispose: () => {
      if (session) close(session)
      ipc.removeHandler(C.connect)
      ipc.removeHandler(C.request)
    },
  }
}
