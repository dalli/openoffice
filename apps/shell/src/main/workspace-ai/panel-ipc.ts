import { WorkspaceDrive, type WorkspaceDriveServices } from './drive'
import { randomUUID } from 'node:crypto'
import type { SkillHandoff } from '@genoffice/electron-utils/generate-handoff'
import type { UserSkillTemplateOpenInput } from '@genoffice/electron-utils/user-skills-channels'
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
import { WorkspaceProjectPreparation } from './project-preparation'
import { WorkspaceAttachments, type WorkspaceAttachmentServices } from './attachments'
import type { ProjectFolderServices } from '../projects/project-ipc'
import { WorkspaceStreams, type WorkspaceStreamService } from './streams'
import {
  DOCS_WORKSPACE_AI_ENTRY,
  isDocsWorkspaceAiEntry,
  type DocsWorkspaceAiEntryResult,
} from '../../../../docs/src/shared/workspace-ai-entry'

export interface WorkspacePanelServices {
  openTemplate?(
    input: UserSkillTemplateOpenInput,
    projectId: string | null,
    targetTabId?: string,
  ): Promise<unknown>
  window: BrowserWindow
  tabs: TabManager
  history: WorkspaceConversationHistory
  sources(): Promise<WorkspacePanelSource[]>
  /** Main derives the saved path and editor from the bound target, never from a message. */
  sender(tabId?: string): ChatSender
  projects?: Pick<ProjectFolderServices, 'guide' | 'tools' | 'scopes'>
  projectMemoryChanged?(projectId: string): void
  attachments?: WorkspaceAttachmentServices
  drive?: WorkspaceDriveServices
  stream?: WorkspaceStreamService
}

interface Session {
  entriesReady: boolean
  entries: Map<
    string,
    {
      accept(accepted: boolean, conversationId?: string): void
      current(conversationId?: string): Promise<boolean>
    }
  >
  info: WorkspacePanelSession
  services: WorkspacePanelServices
  execution: WorkspaceExecutionController
  frame: IpcMainInvokeEvent['senderFrame']
  opened: Set<string>
  origins: Map<string, WorkspacePanelSource['origin']>
  targets: Map<string, string>
  rebindings: Map<string, WorkspacePanelSource['origin']>
  writing: Set<string>
  projects: WorkspaceProjectPreparation | null
  attachments: WorkspaceAttachments | null
  drive: WorkspaceDrive | null
  streams: WorkspaceStreams | null
  created: Map<string, WorkspacePanelSource>
  templateCopies: Map<string, { projectId: string; conversationId: string }>
  templateRuns: Map<number, string>
  off(): void
}

/** Private shell bridge. File access is limited to explicit, run-bound attachment actions. */
export function registerWorkspacePanelIpc(
  ipc: Pick<IpcMain, 'handle' | 'removeHandler'>,
  servicesFor: (event: IpcMainInvokeEvent) => WorkspacePanelServices | null,
): {
  offerTemplate(
    event: IpcMainInvokeEvent,
    handoff: SkillHandoff,
    projectId?: string,
  ): Promise<boolean>
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
    for (const entry of s.entries.values()) entry.accept(false)
    s.off()
    s.services.history.releaseMemory()
    s.projects?.dispose()
    s.attachments?.dispose()
    s.drive?.dispose()
    s.streams?.dispose()
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
    const available = await sources(s)
    live(s)
    origin = resolveOrigin(s, origin)
    const key = workspaceConversationKey(origin)
    if (!available.some((item) => workspaceConversationKey(item.origin) === key))
      fail('invalid-request')
    return { key, requestedKey, origin }
  }
  const sources = async (s: Session): Promise<WorkspacePanelSource[]> => {
    const available = await s.services.sources()
    if (s.created.size && s.services.projects) {
      const listed = await s.services.projects.guide.listProjects()
      for (const item of s.created.values()) {
        const origin = item.origin
        if (
          origin.kind === 'project' &&
          listed.some((p) => p.id === origin.projectId && !p.missing) &&
          !available.some(
            (source) =>
              workspaceConversationKey(source.origin) === workspaceConversationKey(origin),
          )
        )
          available.push(item)
      }
    }
    return [
      ...new Map(available.map((item) => [workspaceConversationKey(item.origin), item])).values(),
    ]
  }
  const catalog = async (s: Session): Promise<WorkspacePanelCatalog> => {
    const available = await sources(s)
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
    return { documents, sources: available }
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
        onQueueChanged: (tabId) => {
          if (session?.info === info)
            emit(session, { sessionId: info.id, kind: 'edit-queue-changed', tabId })
        },
        onStopped: (run) => {
          if (session?.info === info) {
            session.templateRuns.delete(run.bindingId)
            session.projects?.release({ ...run, owner: info.owner })
            session.attachments?.release({ ...run, owner: info.owner })
            session.drive?.release({ ...run, owner: info.owner })
            emit(session, { sessionId: info.id, kind: 'stopped', run })
          }
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
        entriesReady: false,
        entries: new Map(),
        info,
        services,
        execution,
        frame: event.senderFrame,
        opened: new Set(),
        origins: new Map(),
        targets: new Map(),
        rebindings: new Map(),
        writing: new Set(),
        created: new Map(),
        templateCopies: new Map(),
        templateRuns: new Map(),
        drive: services.drive ? new WorkspaceDrive(services.drive) : null,
        streams: services.stream ? new WorkspaceStreams(services.stream) : null,
        attachments: services.attachments ? new WorkspaceAttachments(services.attachments) : null,
        projects: services.projects
          ? new WorkspaceProjectPreparation(
              services.projects,
              event.sender.id,
              services.projectMemoryChanged,
            )
          : null,
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
      case 'open-template': {
        if (!s.opened.has(r.conversationId)) fail('invalid-request')
        const origin = s.origins.get(r.conversationId) ?? fail('invalid-request')
        await source(s, origin)
        live(s)
        if (
          r.targetTabId &&
          !s.services.tabs
            .workspaceDocuments()
            .some((tab) => tab.id === r.targetTabId && tab.available)
        )
          fail('invalid-request')
        return (s.services.openTemplate ?? fail('invalid-request'))(
          r.input,
          origin.kind === 'project' ? origin.projectId : null,
          r.targetTabId,
        )
      }
      case 'entry-ack': {
        const entry = s.entries.get(r.entryId)
        if (!entry) return false
        let accepted = false
        try {
          accepted = r.accepted && (await entry.current(r.conversationId))
        } catch {
          // A replaced/closed document keeps its original queue.
        }
        if (session !== s || s.entries.get(r.entryId) !== entry) return false
        entry.accept(accepted, r.conversationId)
        return accepted
      }
      case 'entries-ready':
        s.entriesReady = r.ready
        if (!r.ready) for (const entry of s.entries.values()) entry.accept(false)
        return
      case 'edit-queue':
        return e.editQueue(r.document, r.command)
      case 'template':
        return e.template(r.document, r.mode)
      case 'stream': {
        if (e.state(r.run.runId)?.status !== 'running') fail('invalid-request')
        s.drive?.check(r.run)
        return (s.streams ?? fail('invalid-request')).stream(r.run, r.request, (chunk) =>
          emit(s, { sessionId: s.info.id, kind: 'stream', run: r.run, chunk }),
        )
      }
      case 'stream-cancel':
        return (s.streams ?? fail('invalid-request')).cancel(r.run, r.requestId)
      case 'drive-status':
        return s.services.drive?.status() ?? { connected: false, enabled: false, allowed: true }
      case 'drive-pick':
        if (!s.opened.has(r.conversationId)) fail('invalid-request')
        return s.services.drive?.pick() ?? { ok: false, reason: 'not-connected' }
      case 'drive-check':
      case 'drive-search':
      case 'drive-read': {
        if (e.state(r.run.runId)?.status !== 'running') fail('invalid-request')
        if (!s.drive && r.kind === 'drive-check')
          return { connected: false, enabled: false, allowed: true }
        const drive = s.drive ?? fail('invalid-request')
        if (r.kind === 'drive-check') return drive.check(r.run)
        return r.kind === 'drive-search'
          ? drive.tool(r.run, 'search', r.query, r.limit)
          : drive.tool(r.run, 'read', r.fileId, r.offset)
      }
      case 'attachment-paste': {
        if (!s.opened.has(r.conversationId)) fail('invalid-request')
        return (s.attachments ?? fail('invalid-request')).paste(r.conversationId, r.bytes, r.ext)
      }
      case 'attachment-pick':
      case 'attachment-add': {
        if (!s.opened.has(r.conversationId)) fail('invalid-request')
        return (s.attachments ?? fail('invalid-request')).add(
          r.conversationId,
          r.kind === 'attachment-add' ? r.paths : undefined,
        )
      }
      case 'attachment-list':
      case 'attachment-read':
      case 'attachment-image': {
        if (e.state(r.run.runId)?.status !== 'running') fail('invalid-request')
        const attachments = s.attachments ?? fail('invalid-request')
        if (r.kind === 'attachment-list') return attachments.list(r.run)
        if (r.kind === 'attachment-image') return attachments.image(r.run, r.index)
        return attachments.read(r.run, r.index, r.offset)
      }
      case 'catalog':
        return catalog(s)
      case 'layout':
        return s.services.tabs.setWorkspacePanel(r.panel)
      case 'capture':
        return e.capture(r.tabId)
      case 'refresh':
        return e.refresh(r.captureId)
      case 'reference': {
        const { id, binding } = await e.readReference(r.captureId, { scope: r.scope })
        return { id, binding }
      }
      case 'create': {
        if (!r.projectId) return { origin: s.services.history.createLocal(), title: '' }
        const project =
          (await s.services.projects?.guide.listProjects())?.find(
            (p) => p.id === r.projectId && !p.missing,
          ) ?? fail('invalid-request')
        live(s)
        const origin = await s.services.history.createProject(r.projectId)
        live(s)
        const source = { origin, title: '', projectTitle: project.name }
        s.created.set(workspaceConversationKey(origin), source)
        return source
      }
      case 'project-prepare': {
        if (!s.opened.has(r.conversationId)) fail('invalid-request')
        const origin = s.origins.get(r.conversationId) ?? fail('invalid-request')
        const ref = await source(s, origin)
        if (ref.origin.kind !== 'project' || ref.origin.projectId !== r.projectId)
          fail('invalid-request')
        return (s.projects ?? fail('invalid-request')).prepare(r.conversationId, r.projectId)
      }
      case 'project-discard':
        s.projects?.discard(r.conversationId, r.binding)
        return
      case 'project-context': {
        const state = e.state(r.run.runId)
        if (state?.status !== 'running') fail('invalid-request')
        return (s.projects ?? fail('invalid-request')).context(r.run)
      }
      case 'project-tool': {
        const state = e.state(r.run.runId)
        if (state?.status !== 'running') fail('invalid-request')
        return (s.projects ?? fail('invalid-request')).tool(r.run, r.method, r.args)
      }
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
        s.drive?.restore(key, history.messages)
        s.drive?.restore(requestedKey, history.messages)
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
        try {
          if (s.attachments) await s.attachments.bind(binding, r.attachments ?? [])
          else if (r.attachments?.length) fail('invalid-request')
          live(s)
          if (binding.project) (s.projects ?? fail('invalid-request')).bind(binding)
          else if (origin.kind === 'project') fail('invalid-request')
          const copy = s.templateCopies.get(r.request.targetId)
          if (copy) {
            const { projectId } = copy
            if (
              copy.conversationId !== r.request.conversationId ||
              origin.kind !== 'project' ||
              origin.projectId !== projectId ||
              binding.project?.id !== projectId
            )
              fail('invalid-request')
            s.templateCopies.delete(r.request.targetId)
            s.templateRuns.set(binding.bindingId, projectId)
          }
        } catch (error) {
          await e.cancel(binding)
          throw error
        }
        return binding
      }
      case 'startNext': {
        await s.streams?.settled()
        live(s)
        const run = await e.startNext()
        if (run) {
          try {
            const origin = s.origins.get(run.binding.conversationId) ?? fail('invalid-request')
            const { key } = await source(s, origin)
            s.targets.set(key, run.binding.target.document.tabId)
            s.targets.set(run.binding.conversationId, run.binding.target.document.tabId)
            await s.projects?.start(run.binding)
            if (s.drive) {
              const attachments = (await s.attachments?.list(run.binding)) ?? []
              s.drive.bind(run, [
                ...[run.binding.target, ...run.binding.references].flatMap((doc) => {
                  const path = s.services.sender(doc.document.tabId).docPath
                  return path ? [path] : []
                }),
                ...attachments.map((file) => file.path),
              ])
              s.drive.check(run.binding)
            }
            live(s)
            s.streams?.bind(run)
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
      case 'writer-update':
        return e.updateWriter(...r.args)
      case 'execute':
        return e.execute(...r.args)
      case 'verifyResponse':
        return e.verifyResponse(...r.args)
      case 'complete': {
        await s.streams?.release(r.args[0])
        const projectId = s.templateRuns.get(r.args[0].bindingId)
        s.templateRuns.delete(r.args[0].bindingId)
        try {
          if (projectId) await e.completeProjectCopy(r.args[0], projectId)
          else await e.complete(...r.args)
        } finally {
          s.projects?.release(r.args[0])
          s.attachments?.release(r.args[0])
          s.drive?.release(r.args[0])
        }
        return
      }
      case 'cancel':
        await e.cancel(...r.args)
        await s.streams?.release(r.args[0])
        return
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
  ipc.handle(
    DOCS_WORKSPACE_AI_ENTRY,
    async (event, raw: unknown): Promise<DocsWorkspaceAiEntryResult> => {
      const s = session
      if (
        !s?.entriesReady ||
        !isDocsWorkspaceAiEntry(raw) ||
        event.sender.isDestroyed() ||
        event.senderFrame !== event.sender.mainFrame ||
        s.services.window.isDestroyed()
      )
        return false
      const tabs = s.services.tabs
      const id = tabs.tabIdForWebContents(event.sender.id)
      if (
        !id ||
        tabs.webContentsForTab(id) !== event.sender ||
        !tabs
          .workspaceDocuments()
          .some((tab) => tab.id === id && tab.kind === 'docs' && tab.available)
      )
        return false
      try {
        // Capture the source view before asynchronous conversation creation or focus changes.
        const capture =
          raw.kind === 'edit-batch'
            ? await s.execution.captureEditBatch(id, raw)
            : await s.execution.capture(id)
        if (
          session !== s ||
          !s.entriesReady ||
          event.sender.isDestroyed() ||
          event.senderFrame !== event.sender.mainFrame ||
          tabs.webContentsForTab(id) !== event.sender
        )
          return raw.kind === 'edit-batch' ? 'rejected' : false
        if (raw.kind === 'edit-batch') {
          if (s.entries.size >= 32) return 'rejected'
          const entryId = randomUUID()
          return await new Promise<DocsWorkspaceAiEntryResult>((resolve) => {
            const timer = setTimeout(() => accept(false), 30_000)
            const accept = (accepted: boolean) => {
              clearTimeout(timer)
              s.entries.delete(entryId)
              resolve(accepted ? true : 'rejected')
            }
            s.entries.set(entryId, {
              accept,
              current: async () => {
                if (
                  session !== s ||
                  !s.entriesReady ||
                  event.sender.isDestroyed() ||
                  event.senderFrame !== event.sender.mainFrame ||
                  tabs.webContentsForTab(id) !== event.sender
                )
                  return false
                const now = await s.execution.captureEditBatch(id, raw)
                return (
                  tabs.webContentsForTab(id) === event.sender &&
                  now.binding.revision === capture.binding.revision &&
                  now.binding.document.documentId === capture.binding.document.documentId
                )
              },
            })
            emit(s, { sessionId: s.info.id, kind: 'entry', entry: raw, capture, entryId })
          })
        }
        emit(s, { sessionId: s.info.id, kind: 'entry', entry: raw, capture })
        return true
      } catch {
        return raw.kind === 'edit-batch' ? 'rejected' : false
      }
    },
  )
  return {
    async offerTemplate(event, handoff, projectId) {
      const s = session
      if (
        !s?.entriesReady ||
        handoff.target !== 'docs' ||
        handoff.preview ||
        event.sender.isDestroyed() ||
        event.senderFrame !== event.sender.mainFrame ||
        s.services.window.isDestroyed() ||
        s.entries.size >= 32
      )
        return false
      const tabs = s.services.tabs
      const id = tabs.tabIdForWebContents(event.sender.id)
      const current = () =>
        session === s &&
        s.entriesReady &&
        !event.sender.isDestroyed() &&
        event.senderFrame === event.sender.mainFrame &&
        !!id &&
        tabs.webContentsForTab(id) === event.sender
      if (
        !current() ||
        !tabs
          .workspaceDocuments()
          .some((tab) => tab.id === id && tab.kind === 'docs' && tab.available)
      )
        return false
      try {
        const capture = await s.execution.capture(id!)
        if (!current()) return false
        const entryId = randomUUID()
        return await new Promise<boolean>((resolve) => {
          const timer = setTimeout(() => accept(false), 30_000)
          const accept = (accepted: boolean, conversationId?: string) => {
            clearTimeout(timer)
            s.entries.delete(entryId)
            if (accepted && projectId && conversationId) {
              s.templateCopies.set(capture.id, { projectId, conversationId })
              if (s.templateCopies.size > 128)
                s.templateCopies.delete(s.templateCopies.keys().next().value!)
            }
            resolve(accepted)
          }
          s.entries.set(entryId, {
            accept,
            current: async (conversationId) => {
              if (!current()) return false
              if (projectId) {
                const origin = conversationId ? s.origins.get(conversationId) : null
                if (
                  !conversationId ||
                  !s.opened.has(conversationId) ||
                  origin?.kind !== 'project' ||
                  origin.projectId !== projectId
                )
                  return false
              }
              const now = await s.execution.refresh(capture.id)
              return (
                current() &&
                now.binding.revision === capture.binding.revision &&
                now.binding.document.documentId === capture.binding.document.documentId
              )
            },
          })
          try {
            emit(s, {
              sessionId: s.info.id,
              kind: 'entry',
              capture,
              entryId,
              entry: {
                kind: 'template-handoff',
                id: handoff.id,
                text: handoff.text,
                chips: handoff.chips,
                ...(projectId ? { projectId } : {}),
              },
            })
          } catch {
            accept(false)
          }
        })
      } catch {
        return false
      }
    },
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
      ipc.removeHandler(DOCS_WORKSPACE_AI_ENTRY)
    },
  }
}
