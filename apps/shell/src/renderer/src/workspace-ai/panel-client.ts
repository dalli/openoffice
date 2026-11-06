import type { WorkspaceReadRequest } from '@genoffice/agent-core'
import type { WorkspaceEditQueueCommand } from '@genoffice/agent-core/workspace-editor-protocol'
import type { WorkspaceDocumentIdentity } from '../../../shared/workspace-ai-api'
import type { WorkspaceConversationOrigin } from '../../../shared/workspace-ai-api'
import type { WorkspaceHistoryMessage } from '../../../main/workspace-ai/conversation-history'
import type {
  WorkspacePanelApi,
  WorkspacePanelEvent,
  WorkspacePanelReply,
  WorkspacePanelRequest,
  WorkspacePanelSession,
  WorkspacePanelValues,
} from '../../../shared/workspace-panel-api'
import type { WorkspacePanelLayout } from '../../../shared/workspace-layout'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import {
  WorkspaceRunHost,
  type WorkspaceExecutionBridge,
  type WorkspaceRunHostDeps,
  type WorkspaceRunPreparation,
} from './run-host'
import { WorkspacePanelController } from './panel-controller'
import type { Lang } from '@genoffice/i18n'
import { aiSkillsBridge, skillTemplateBridge } from '@genoffice/ui'
import {
  windowUserSkillsBridge,
  windowMcpToolsBridge,
  windowOutboundBridge,
  type McpToolsBridgeLike,
  type OutboundBridgeLike,
} from '@genoffice/agent-core'
import type { WorkspaceProjectBinding } from '../../../shared/workspace-ai-api'
import type { WorkspaceRunHandle } from '../../../main/workspace-ai/execution-controller'
import type { WorkspaceProjectTool } from '../../../shared/workspace-project-api'
import { prepareWorkspaceAttachments } from './attachment-preparation'
import { prepareWorkspaceProject } from './project-preparation'
import { WorkspaceUserSkills, type WorkspaceSkillsBridge } from './user-skills'
import { WorkspaceUserMcp } from './user-mcp'
import { WorkspaceDrivePreparation } from './drive-preparation'

function unwrap<T>(reply: WorkspacePanelReply<T>): T {
  if (!reply.ok) throw new Error(`Workspace panel rejected: ${reply.reason}`)
  return reply.value
}

/** Restores process-local abort signals; no DOM or provider dependency. */
export class WorkspacePanelClient implements WorkspaceExecutionBridge {
  private readonly signals = new Map<number, AbortController>()
  private readonly stopped = new Set<number>()
  private readonly listeners = new Set<(event: WorkspacePanelEvent) => void>()
  private closed = false
  private off: () => void = () => {}
  private constructor(
    private readonly api: WorkspacePanelApi,
    readonly session: WorkspacePanelSession,
  ) {}

  static async connect(api: WorkspacePanelApi): Promise<WorkspacePanelClient> {
    const client = new WorkspacePanelClient(api, unwrap(await api.connect()))
    client.off = api.onEvent((event) => client.receive(event))
    return client
  }
  private receive(event: WorkspacePanelEvent): void {
    if (this.closed || event.sessionId !== this.session.id) return
    if (event.kind === 'stopped') {
      this.stopped.add(event.run.bindingId)
      if (this.stopped.size > 128) this.stopped.delete(this.stopped.values().next().value!)
      this.signals.get(event.run.bindingId)?.abort()
      this.signals.delete(event.run.bindingId)
    }
    if (event.kind === 'closed') this.invalidate()
    for (const listener of [...this.listeners]) listener(event)
  }
  private invalidate(): void {
    this.closed = true
    for (const signal of this.signals.values()) signal.abort()
    this.signals.clear()
    this.stopped.clear()
    this.off()
  }
  subscribe(listener: (event: WorkspacePanelEvent) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  private async call<R extends WorkspacePanelRequest>(
    request: R,
  ): Promise<WorkspacePanelValues[R['kind']]> {
    if (this.closed) throw new Error('Workspace panel closed')
    const reply = await this.api.request(this.session.id, request)
    if (this.closed) throw new Error('Workspace panel closed')
    return unwrap(reply)
  }
  entriesReady = (ready: boolean) => this.call({ kind: 'entries-ready', ready })
  editQueue = (document: WorkspaceDocumentIdentity, command: WorkspaceEditQueueCommand) =>
    this.call({ kind: 'edit-queue', document, command })
  acknowledgeEntry = (entryId: string, accepted: boolean, conversationId?: string) =>
    this.call({ kind: 'entry-ack', entryId, accepted, conversationId })
  driveStatus = () => this.call({ kind: 'drive-status' })
  stream = (
    run: WorkspaceRunHandle,
    request: import('../../../main/workspace-ai/streams').WorkspaceStreamRequest,
  ) => this.call({ kind: 'stream', run, request })
  cancelStream = (run: WorkspaceRunHandle, requestId: string) =>
    this.call({ kind: 'stream-cancel', run, requestId })
  pickDriveAttachments = (conversationId: string) =>
    this.call({ kind: 'drive-pick', conversationId })
  checkDrive = (run: WorkspaceRunHandle) => this.call({ kind: 'drive-check', run })
  searchDrive = (run: WorkspaceRunHandle, query: string, limit: number) =>
    this.call({ kind: 'drive-search', run, query, limit })
  readDrive = (run: WorkspaceRunHandle, fileId: string, offset: number) =>
    this.call({ kind: 'drive-read', run, fileId, offset })
  pasteAttachment = (conversationId: string, bytes: ArrayBuffer, ext: string) =>
    this.call({ kind: 'attachment-paste', conversationId, bytes, ext })
  pickAttachments = (conversationId: string) =>
    this.call({ kind: 'attachment-pick', conversationId })
  addAttachments = (conversationId: string, paths: string[]) =>
    this.call({ kind: 'attachment-add', conversationId, paths })
  attachmentList = (run: WorkspaceRunHandle) => this.call({ kind: 'attachment-list', run })
  attachmentRead = (run: WorkspaceRunHandle, index: number, offset: number) =>
    this.call({ kind: 'attachment-read', run, index, offset })
  attachmentImage = (run: WorkspaceRunHandle, index: number) =>
    this.call({ kind: 'attachment-image', run, index })
  catalog = () => this.call({ kind: 'catalog' })
  template = (
    document: import('../../../shared/workspace-ai-api').WorkspaceDocumentIdentity,
    mode: import('@genoffice/agent-core').WorkspaceTemplateMode,
  ) => this.call({ kind: 'template', document, mode })
  layout = (panel: WorkspacePanelLayout | null) => this.call({ kind: 'layout', panel })
  capture = (tabId: string) => this.call({ kind: 'capture', tabId })
  refresh = (captureId: string) => this.call({ kind: 'refresh', captureId })
  reference = (captureId: string, scope: WorkspaceReadRequest['scope']) =>
    this.call({ kind: 'reference', captureId, scope })
  create = (projectId?: string) => this.call({ kind: 'create', projectId })
  openTemplate = (
    conversationId: string,
    input: Extract<WorkspacePanelRequest, { kind: 'open-template' }>['input'],
    targetTabId?: string,
  ) => this.call({ kind: 'open-template', conversationId, input, targetTabId })
  prepareProject = (conversationId: string, projectId: string) =>
    this.call({ kind: 'project-prepare', conversationId, projectId })
  discardProject = (conversationId: string, binding: WorkspaceProjectBinding) =>
    this.call({ kind: 'project-discard', conversationId, binding })
  projectContext = (run: WorkspaceRunHandle) => this.call({ kind: 'project-context', run })
  projectTool = (
    run: WorkspaceRunHandle,
    method: WorkspaceProjectTool,
    args: Record<string, unknown>,
  ) => this.call({ kind: 'project-tool', run, method, args })
  load = (origin: WorkspaceConversationOrigin) => this.call({ kind: 'load', origin })
  append = (origin: WorkspaceConversationOrigin, message: WorkspaceHistoryMessage) =>
    this.call({ kind: 'append', origin, message })
  enqueue: WorkspaceExecutionBridge['enqueue'] = (request, attachments) =>
    this.call({ kind: 'enqueue', request, attachments })
  async startNext(): Promise<WorkspaceStartedRun | null> {
    const run = await this.call({ kind: 'startNext' })
    if (!run) return null
    const abort = new AbortController()
    const id = run.binding.bindingId
    if (run.aborted || this.stopped.has(id)) abort.abort()
    else this.signals.set(id, abort)
    return {
      binding: run.binding,
      target: run.target,
      references: run.references,
      signal: abort.signal,
    }
  }
  updateWriter: NonNullable<WorkspaceExecutionBridge['updateWriter']> = (...args) =>
    this.call({ kind: 'writer-update', args })
  execute: WorkspaceExecutionBridge['execute'] = (...args) => this.call({ kind: 'execute', args })
  verifyResponse: WorkspaceExecutionBridge['verifyResponse'] = (...args) =>
    this.call({ kind: 'verifyResponse', args })
  complete: WorkspaceExecutionBridge['complete'] = async (...args) => {
    await this.call({ kind: 'complete', args })
    this.signals.delete(args[0].bindingId)
  }
  cancel: WorkspaceExecutionBridge['cancel'] = async (...args) => {
    // Main still validates the complete owner/run handle before touching any editor.
    await this.call({ kind: 'cancel', args })
  }
  state: WorkspaceExecutionBridge['state'] = (runId) => this.call({ kind: 'state', runId })
  rollbackForRun: WorkspaceExecutionBridge['rollbackForRun'] = (...args) =>
    this.call({ kind: 'rollbackForRun', args })
  rollback: WorkspaceExecutionBridge['rollback'] = (...args) =>
    this.call({ kind: 'rollback', args })
  async dispose(): Promise<void> {
    if (this.closed) return
    this.invalidate()
    this.listeners.clear()
    unwrap(await this.api.request(this.session.id, { kind: 'disconnect' }))
  }
}

/** The caller must supply the existing gate, guarded preparation and transport explicitly. */
export async function connectWorkspacePanel(
  api: WorkspacePanelApi,
  deps: Omit<WorkspaceRunHostDeps, 'execution' | 'load' | 'append' | 'createTransport'> & {
    createTransport(
      id: string,
      client: WorkspacePanelClient,
      purpose?: 'writer',
    ): ReturnType<WorkspaceRunHostDeps['createTransport']>
    lang?: () => Lang
    userSkillsBridge?: WorkspaceSkillsBridge | null
    mcpBridge?: McpToolsBridgeLike | null
    outboundBridge?: OutboundBridgeLike | null
  },
) {
  const client = await WorkspacePanelClient.connect(api)
  let userSkills: WorkspaceUserSkills | undefined
  let userMcp: WorkspaceUserMcp | undefined
  const host: WorkspaceRunHost = new WorkspaceRunHost({
    ...deps,
    execution: client,
    load: client.load,
    append: client.append,
    createTransport: (id, purpose) => deps.createTransport(id, client, purpose),
    captureProject: async (request) => {
      const origin = host.conversations.state(request.conversationId).origin
      if (origin.kind !== 'project') return null
      const prepared = await client.prepareProject(request.conversationId, origin.projectId)
      controller.projectNotices(request.conversationId, prepared.notices)
      return prepared.binding
    },
    discardProject: client.discardProject,
    prepare: async (input, run, permissions): Promise<WorkspaceRunPreparation> => {
      const project = await prepareWorkspaceProject(client, run, deps.lang ?? (() => 'en'))
      let attachments: WorkspaceRunPreparation | undefined
      let skills: WorkspaceRunPreparation | undefined
      let mcp: WorkspaceRunPreparation | undefined
      let preparedDrive: WorkspaceRunPreparation | undefined
      try {
        attachments = await prepareWorkspaceAttachments(
          client,
          input,
          run,
          deps.lang ?? (() => 'en'),
        )
        preparedDrive = await drive.prepare(run)
        mcp = await userMcp?.prepare(run)
        skills = await userSkills?.prepare(input, run)
        if (!skills && input.skills.length) throw new Error('Skills unavailable')
        const other = await deps.prepare(input, run, permissions)
        if (
          other.project &&
          (other.project.id !== project.project?.id ||
            other.project.contextVersion !== project.project?.contextVersion)
        ) {
          other.dispose()
          throw new Error('Project context changed')
        }
        return {
          ...other,
          project: project.project,
          skills: [
            ...project.skills,
            ...attachments.skills,
            ...preparedDrive.skills,
            ...(skills?.skills ?? []),
            ...(mcp?.skills ?? []),
            ...other.skills,
          ],
          guard: (skill) => {
            const guarded = other.guard?.(skill) ?? skill
            const withDrive = preparedDrive?.guard?.(guarded) ?? guarded
            return mcp?.guard?.(withDrive) ?? withDrive
          },
          userText:
            (other.userText ?? input.text) + (skills?.userText?.slice(input.text.length) ?? ''),
          async beforeTurn() {
            await preparedDrive?.beforeTurn?.()
            await other.beforeTurn?.()
          },
          onStart() {
            skills?.onStart?.()
            other.onStart?.()
          },
          images: [...(attachments.images ?? []), ...(other.images ?? [])],
          dispose() {
            try {
              other.dispose()
            } finally {
              attachments?.dispose()
              skills?.dispose()
              mcp?.dispose()
              preparedDrive?.dispose()
              project.dispose()
            }
          },
        }
      } catch (error) {
        attachments?.dispose()
        skills?.dispose()
        mcp?.dispose()
        preparedDrive?.dispose()
        project.dispose()
        throw error
      }
    },
  })
  const drive: WorkspaceDrivePreparation = new WorkspaceDrivePreparation(client, host)
  const skillBridge =
    deps.userSkillsBridge !== undefined
      ? deps.userSkillsBridge
      : aiSkillsBridge() && windowUserSkillsBridge()
        ? (aiSkillsBridge() as WorkspaceSkillsBridge)
        : null
  if (skillBridge)
    userSkills = new WorkspaceUserSkills(
      skillBridge,
      host,
      deps.lang ?? (() => 'en'),
      skillTemplateBridge(),
      (id, server) => userMcp?.usableServers(id).has(server) ?? false,
      client.template,
      (id, input) =>
        client.openTemplate(
          id,
          input,
          host.conversations.state(id).composer.target?.document.tabId,
        ),
    )
  const mcpBridge = deps.mcpBridge === undefined ? windowMcpToolsBridge() : deps.mcpBridge
  const outboundBridge =
    deps.outboundBridge === undefined ? windowOutboundBridge() : deps.outboundBridge
  if (mcpBridge && outboundBridge)
    userMcp = new WorkspaceUserMcp(mcpBridge, outboundBridge, host, deps.lang ?? (() => 'en'))
  const controller = new WorkspacePanelController({
    handoff: (id, capture, entry) => userSkills?.acceptHandoff(id, capture, entry) ?? false,
    host,
    capture: client.capture,
    refresh: client.refresh,
    reference: client.reference,
    create: client.create,
  })
  let closed = false
  let epoch = 0
  const refresh = async () => {
    const current = ++epoch
    const catalog = await client.catalog()
    if (!closed && epoch === current) controller.catalog(catalog.documents, catalog.sources)
  }
  const off = client.subscribe((event) => {
    if (event.kind === 'history-changed') host.historyChanged(event.origin)
    if (event.kind === 'history-rebound') host.rebindHistory(event.previous, event.origin)
    if (event.kind === 'catalog-changed') void refresh().catch(() => {})
    if (event.kind === 'closed') {
      closed = true
      controller.dispose()
      host.dispose()
      userSkills?.dispose()
      userMcp?.dispose()
      drive.dispose()
      off()
    }
  })
  try {
    await refresh()
  } catch (error) {
    off()
    controller.dispose()
    host.dispose()
    userSkills?.dispose()
    userMcp?.dispose()
    drive.dispose()
    await client.dispose().catch(() => {})
    throw error
  }
  return {
    client,
    host,
    controller,
    userSkills,
    userMcp,
    refresh,
    async dispose() {
      closed = true
      off()
      controller.dispose()
      host.dispose()
      userSkills?.dispose()
      userMcp?.dispose()
      drive.dispose()
      await client.dispose()
    },
  }
}
