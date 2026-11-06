import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react'
import type { Lang } from '@genoffice/i18n'
import type {
  ProjectFolderApi,
  ProjectListItem,
  AddFolderResult,
  CreateProjectResult,
} from '@genoffice/project-store'
import {
  ProjectPicker,
  ProjectProblemLine,
  projectAddProblem,
  type ProjectProblem,
} from '@genoffice/ui'
import type { WorkspacePanelController } from './panel-controller'

/** Existing project management/trust dialogs; the choice starts a separate workspace conversation. */
export function WorkspaceProjectControls({
  controller,
  bridge,
  lang,
}: {
  controller: WorkspacePanelController
  bridge: ProjectFolderApi
  lang: Lang
}) {
  useSyncExternalStore(controller.subscribe, controller.version, controller.version)
  const [projects, setProjects] = useState<ProjectListItem[]>([])
  const [parent, setParent] = useState<string>()
  const [problem, setProblem] = useState<ProjectProblem | null>(null)
  const [busy, setBusy] = useState(false)
  const mounted = useRef(false)
  const epoch = useRef(0)
  const id = controller.host.conversations.selectedId
  const state = id ? controller.host.conversations.state(id) : null
  const projectId = state?.origin.kind === 'project' ? state.origin.projectId : null
  const refresh = useCallback(async () => {
    const mine = ++epoch.current
    try {
      const [list, folder] = await Promise.all([
        bridge.listProjects(),
        bridge.getDefaultParentFolder(),
      ])
      if (!mounted.current || mine !== epoch.current) return
      setProjects(list)
      setParent(folder)
    } catch {
      if (mounted.current && mine === epoch.current) setProblem({ kind: 'action-failed' })
    }
  }, [bridge])
  useEffect(() => {
    mounted.current = true
    void refresh()
    return () => {
      mounted.current = false
    }
  }, [refresh])
  const action = async (run: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    setProblem(null)
    try {
      await run()
    } catch {
      if (mounted.current) setProblem({ kind: 'action-failed' })
    } finally {
      if (mounted.current) setBusy(false)
    }
  }
  const adopt = async (result: AddFolderResult | CreateProjectResult) => {
    if (!mounted.current) return
    const refused = projectAddProblem(result)
    if (refused) {
      setProblem(refused)
      return
    }
    if (result.status === 'added' || result.status === 'created') {
      await refresh()
      if (mounted.current)
        await controller.create(result.project.id, result.status === 'added' ? result.notices : [])
    }
  }
  return (
    <>
      <ProjectPicker
        lang={lang}
        projects={projects}
        selectedId={projectId}
        disabled={busy}
        defaultParentFolder={parent}
        onOpen={() => {
          void refresh()
        }}
        onSelect={(next) => {
          if (next !== projectId) void action(() => controller.create(next ?? undefined))
        }}
        onCreate={(name, options) => {
          void action(async () => adopt(await bridge.createProject({ name, ...options })))
        }}
        onAddFolder={() => {
          void action(async () => adopt(await bridge.addFolder()))
        }}
        onRelocate={(projectId) => {
          void action(async () => adopt(await bridge.relocateProject({ projectId })))
        }}
        onSetPinned={(projectId, pinned) => {
          void action(async () => {
            await bridge.setPinned({ projectId, pinned })
            await refresh()
          })
        }}
        onOpenInFileManager={(projectId) => {
          void action(() => bridge.openInFileManager({ projectId }))
        }}
        onRemoveFromList={(projectId) => {
          void action(async () => {
            await bridge.removeFromList({ projectId })
            await refresh()
          })
        }}
      />
      {problem && <ProjectProblemLine lang={lang} problem={problem} />}
    </>
  )
}
