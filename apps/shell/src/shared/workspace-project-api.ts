import type { ProjectContextBundle, ProjectNotice } from '@genoffice/project-store'
import type { WorkspaceProjectBinding } from './workspace-ai-api'

export interface WorkspacePreparedProject {
  binding: WorkspaceProjectBinding
  notices: ProjectNotice[]
}

export interface WorkspaceProjectContext {
  binding: WorkspaceProjectBinding
  context: ProjectContextBundle
}

/** A deliberately closed set of existing app tools; no raw scope or arbitrary IPC is exposed. */
export type WorkspaceProjectTool = 'listFiles' | 'search' | 'read' | 'remember'
