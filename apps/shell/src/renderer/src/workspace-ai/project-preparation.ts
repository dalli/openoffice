import type { Lang } from '@genoffice/i18n'
import type { ProjectFolderApi } from '@genoffice/project-store'
import { panelProjectSkill, type ProjectRun } from '@genoffice/ui'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import type { WorkspaceProjectTool } from '../../../shared/workspace-project-api'
import type { WorkspacePanelClient } from './panel-client'
import type { WorkspaceRunPreparation } from './run-host'

/** Reuse the existing project prompt and tools with only main-owned run capabilities. */
export async function prepareWorkspaceProject(
  client: Pick<WorkspacePanelClient, 'projectContext' | 'projectTool'>,
  run: WorkspaceStartedRun,
  lang: () => Lang,
): Promise<WorkspaceRunPreparation> {
  if (!run.binding.project) return { project: null, skills: [], dispose() {} }
  const prepared = await client.projectContext(run.binding)
  if (
    !prepared ||
    prepared.binding.id !== run.binding.project.id ||
    prepared.binding.contextVersion !== run.binding.project.contextVersion ||
    run.signal.aborted
  )
    throw new Error('Project preparation is no longer available')
  let closed = false
  let unwritable = false
  const snapshot = { context: prepared.context, scope: run.binding.runId }
  const invoke = async <M extends WorkspaceProjectTool>(
    method: M,
    args: Parameters<ProjectFolderApi[M]>[0],
  ): Promise<Awaited<ReturnType<ProjectFolderApi[M]>>> => {
    if (closed || run.signal.aborted || args.scope !== run.binding.runId)
      throw new Error('Project execution stopped')
    const { scope: _, ...input } = args
    const result = await client.projectTool(run.binding, method, input)
    if (closed || run.signal.aborted) throw new Error('Project execution stopped')
    return result as Awaited<ReturnType<ProjectFolderApi[M]>>
  }
  const project: ProjectRun = {
    bridge: {
      readContext: async () => ({ status: 'ok', context: prepared.context }),
      openScope: async () => run.binding.runId,
      closeScope: async () => {},
      listFiles: (args) => invoke('listFiles', args),
      search: (args) => invoke('search', args),
      read: (args) => invoke('read', args),
      remember: (args) => invoke('remember', args),
    },
    get snapshot() {
      return closed ? null : snapshot
    },
    get unwritable() {
      return unwritable
    },
    prepare: async () => ({
      status: 'ok',
      context: prepared.context,
      notices: prepared.context.notices,
    }),
    finish: async () => {
      closed = true
    },
    markUnwritable: () => {
      unwritable = true
    },
  }
  return {
    project: prepared.binding,
    skills: panelProjectSkill({ run: project, lang }),
    dispose: () => {
      closed = true
    },
  }
}
