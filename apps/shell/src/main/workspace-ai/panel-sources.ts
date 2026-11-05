import { basename } from 'node:path'
import type { ProjectStore } from '@genoffice/project-store'
import type { WorkspacePanelSource } from '../../shared/workspace-panel-api'
import type { ProjectGuide } from '../projects/project-guide'
import type { ProjectChatStore } from '../projects/chat-store'
import { workspaceConversationKey } from '../../shared/workspace-ai-api'

/** Lists references only. Never resolves a document chat, loads its body or changes project trust. */
export async function workspacePanelSources(
  local: ProjectStore,
  projects: Pick<ProjectGuide, 'listProjects'>,
  chats: Pick<ProjectChatStore, 'listChats'>,
  live: WorkspacePanelSource[] = [],
): Promise<WorkspacePanelSource[]> {
  const sources: WorkspacePanelSource[] = local.listWorkspaceChats().map((chat) => ({
    origin: { kind: 'workspace', conversationId: chat.id },
    title: chat.title,
  }))
  for (const project of local.listProjects()) {
    const titles = new Map(
      (local.getProject(project.id)?.files ?? []).map((path) => [
        local.chatIdForPath(path, project.id),
        basename(path),
      ]),
    )
    for (const chat of local.listChats(project.id)) {
      if (!/^[A-Za-z0-9_-]+$/.test(chat.chatId)) continue
      sources.push({
        origin: { kind: 'legacy-document', projectId: project.id, chatId: chat.chatId },
        title: titles.get(chat.chatId) ?? chat.chatId,
        projectTitle: project.name,
      })
    }
  }
  for (const project of await projects.listProjects()) {
    if (project.missing) continue
    // The existing service filters foreign records using this PC's current trust decision.
    for (const chat of await chats.listChats(project.id)) {
      sources.push({
        origin: { kind: 'project', projectId: project.id, chatId: chat.id },
        title: chat.title,
        projectTitle: project.name,
      })
    }
  }
  return [
    ...new Map([...sources, ...live].map((s) => [workspaceConversationKey(s.origin), s])).values(),
  ]
}
