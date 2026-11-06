import { createHash, randomUUID } from 'node:crypto'
import type { ProjectContextBundle } from '@genoffice/project-store'
import type { ProjectFolderServices } from '../projects/project-ipc'
import type { WorkspaceAiBoundRun, WorkspaceProjectBinding } from '../../shared/workspace-ai-api'
import type {
  WorkspacePreparedProject,
  WorkspaceProjectContext,
  WorkspaceProjectTool,
} from '../../shared/workspace-project-api'
import { WorkspaceExecutionError, type WorkspaceRunHandle } from './execution-controller'

type Services = Pick<ProjectFolderServices, 'guide' | 'tools' | 'scopes'>
interface Prepared {
  conversationId: string
  binding: WorkspaceProjectBinding
  context: ProjectContextBundle
  fingerprint: string
  run: WorkspaceAiBoundRun | null
  scope: string | null
}
function fail(): never {
  throw new WorkspaceExecutionError('invalid-request')
}

/** Main-owned project snapshots. Selection alone never reads context or grants trust. */
export class WorkspaceProjectPreparation {
  private readonly prepared = new Map<string, Prepared>()
  private closed = false

  constructor(
    private readonly services: Services,
    private readonly owner: number,
    private readonly memoryChanged: (projectId: string) => void = () => {},
  ) {}

  private async read(projectId: string) {
    if (this.closed) fail()
    const read = await this.services.guide.readContext(projectId)
    const visibility = await this.services.guide.aiVisibility(projectId)
    if (this.closed || read.status !== 'ok' || !visibility) fail()
    const { notices: _, ...content } = read.context
    // One-time notices change after a read; the actual context, root and visibility must match.
    const fingerprint = createHash('sha256')
      .update(JSON.stringify([content, visibility.root, visibility.ruleKey]))
      .digest('hex')
    return { context: read.context, fingerprint }
  }

  async prepare(conversationId: string, projectId: string): Promise<WorkspacePreparedProject> {
    const snapshot = await this.read(projectId)
    // Bound runs retain their snapshots. Unused admission capabilities are bounded separately.
    const unused = [...this.prepared].filter(([, value]) => !value.run)
    for (const [key] of unused.slice(0, Math.max(0, unused.length - 31))) this.prepared.delete(key)
    const binding = { id: projectId, contextVersion: randomUUID() }
    this.prepared.set(binding.contextVersion, {
      ...snapshot,
      binding,
      conversationId,
      run: null,
      scope: null,
    })
    return { binding: { ...binding }, notices: structuredClone(snapshot.context.notices) }
  }

  bind(run: WorkspaceAiBoundRun): void {
    if (!run.project) return
    const item = this.prepared.get(run.project.contextVersion)
    if (
      this.closed ||
      !item ||
      item.run ||
      item.binding.id !== run.project.id ||
      item.conversationId !== run.conversationId
    )
      fail()
    item.run = structuredClone(run)
  }

  private entry(run: WorkspaceRunHandle): Prepared {
    const item = [...this.prepared.values()].find((item) => {
      const bound = item.run
      return (
        bound?.bindingId === run.bindingId &&
        bound.runId === run.runId &&
        bound.conversationId === run.conversationId &&
        bound.owner.windowId === run.owner.windowId &&
        bound.owner.panelContentsId === run.owner.panelContentsId &&
        bound.owner.panelGeneration === run.owner.panelGeneration
      )
    })
    if (this.closed || !item) fail()
    return item
  }

  async start(run: WorkspaceAiBoundRun): Promise<void> {
    if (!run.project) return
    const item = this.entry(run)
    if (item.scope) fail()
    const current = await this.read(item.binding.id)
    if (this.entry(run) !== item || current.fingerprint !== item.fingerprint) fail()
    const scope = await this.services.tools.openScope(this.owner, { projectId: item.binding.id })
    if (this.closed || this.prepared.get(item.binding.contextVersion) !== item) {
      this.services.tools.closeScope(this.owner, { scope })
      fail()
    }
    item.scope = scope
  }

  context(run: WorkspaceRunHandle): WorkspaceProjectContext {
    const item = this.entry(run)
    if (!item.scope) fail()
    return structuredClone({ binding: item.binding, context: item.context })
  }

  async tool(run: WorkspaceRunHandle, method: WorkspaceProjectTool, args: Record<string, unknown>) {
    const item = this.entry(run)
    if (!item.scope) fail()
    const scoped = { ...args, scope: item.scope }
    let result: unknown
    if (method === 'remember') {
      if (!item.context.canRemember || item.context.trust !== 'use') fail()
      // The existing writer rechecks current trust, memory settings and writability.
      const change = args.change as Parameters<Services['guide']['editMemory']>[1]
      if (this.services.scopes.projectOf(this.owner, scoped) !== item.binding.id) fail()
      result = await this.services.guide.editMemory(item.binding.id, change, 'ai')
      if ((result as { status: string }).status === 'ok') this.memoryChanged(item.binding.id)
    } else result = await this.services.tools[method](this.owner, scoped)
    if (this.entry(run) !== item) fail()
    return result
  }

  release(run: WorkspaceRunHandle): void {
    let item: Prepared
    try {
      item = this.entry(run)
    } catch {
      return
    }
    this.prepared.delete(item.binding.contextVersion)
    if (item.scope) this.services.tools.closeScope(this.owner, { scope: item.scope })
  }

  discard(conversationId: string, binding: WorkspaceProjectBinding): void {
    const item = this.prepared.get(binding.contextVersion)
    if (
      item &&
      !item.run &&
      item.conversationId === conversationId &&
      item.binding.id === binding.id
    )
      this.prepared.delete(binding.contextVersion)
  }

  dispose(): void {
    this.closed = true
    for (const item of this.prepared.values())
      if (item.scope) this.services.tools.closeScope(this.owner, { scope: item.scope })
    this.prepared.clear()
  }
}
