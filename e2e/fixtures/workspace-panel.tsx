import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import { WorkspacePanelController } from '../../apps/shell/src/renderer/src/workspace-ai/panel-controller'
import { WorkspacePanel } from '../../apps/shell/src/renderer/src/workspace-ai/WorkspacePanel'
import {
  WorkspaceRunHost,
  type WorkspaceExecutionBridge,
} from '../../apps/shell/src/renderer/src/workspace-ai/run-host'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { SplitDivider } from '../../apps/shell/src/renderer/src/SplitDivider'
import type { WorkspaceLayout } from '../../apps/shell/src/shared/workspace-layout'
import type { WorkspaceCapture } from '../../apps/shell/src/main/workspace-ai/execution-controller'
import type { ViewportSize } from '../../apps/shell/src/renderer/src/split-view-geometry'
import type { SplitViewState } from '../../apps/shell/src/shared/split-view-api'
import type { Lang } from '@genoffice/i18n'

// This native composition fixture exercises real panel controls, never a provider or document engine.
document.querySelector('aside')!.remove()
const root = createRoot(document.querySelector('#react-divider')!)
let calls = 0
let next = 0
let requestedOpen: boolean | null = null
const forbidden = (): never => {
  calls++
  throw new Error('No model/editor execution is connected in this fixture')
}
const host = new WorkspaceRunHost({
  execution: new Proxy({} as WorkspaceExecutionBridge, { get: () => forbidden }),
  load: async () => ({ messages: [], foreign: false }),
  append: async () => {
    forbidden()
  },
  admit: () => false,
  gate: () => ({ kind: 'checking' }),
  prepare: async () => forbidden(),
  createTransport: () => ({ stream: forbidden }),
  errorText: () => 'Synthetic fixture error',
})
const captures = new Map<string, WorkspaceCapture>()
const controller = new WorkspacePanelController({
  host,
  capture: async (tabId) => {
    const id = crypto.randomUUID()
    const value: WorkspaceCapture = {
      id,
      binding: {
        document: { windowId: 1, tabId, documentId: tabId, kind: 'docs', generation: 1 },
        revision: '1',
        selection: { id: 'synthetic-selection', revision: '1' },
      },
    }
    captures.set(id, value)
    return value
  },
  reference: async (id) => {
    const referenceId = crypto.randomUUID()
    return { id: referenceId, binding: { ...captures.get(id)!.binding, snapshotId: referenceId } }
  },
  create: async () => ({
    origin: { kind: 'workspace', conversationId: `synthetic-${++next}` },
    title: `Conversation ${next}`,
  }),
})
controller.catalog(
  [
    { id: 'A', kind: 'docs', title: 'Synthetic A', available: true, active: true, pane: 'left' },
    { id: 'B', kind: 'markdown', title: 'Synthetic B', available: true, pane: 'right' },
  ],
  [],
)
const ready = controller.create()

Object.assign(window, {
  aiOfficeTabs: { setSplitRatio: () => undefined },
  workspacePanelProbe: {
    calls: () => calls,
    requestedOpen: () => requestedOpen,
    selected: () => host.conversations.selectedId,
    state: () => host.conversations.state(host.conversations.selectedId!),
  },
  async arrange({
    layout,
    viewport,
    orientation,
    ratio,
    uiScale,
  }: {
    layout: WorkspaceLayout
    viewport: ViewportSize
    orientation: SplitViewState['orientation'] | null
    ratio: number
    uiScale: number
  }) {
    await ready
    document.documentElement.style.setProperty('--ui-scale', String(uiScale))
    const split = orientation
      ? { aTabId: 'A', bTabId: 'B', focus: 'b' as const, ratio, orientation }
      : null
    const lang = (document.documentElement.lang || 'en') as Lang
    flushSync(() =>
      root.render(
        <LocaleProvider key={lang} initial={lang}>
          <div className="app-frame">
            {split && (
              <SplitDivider split={split} viewport={{ ...viewport, documents: layout.documents }} />
            )}
          </div>
          <WorkspacePanel
            mount={{
              controller,
              layout,
              uiScale,
              setOpen: (open) => {
                requestedOpen = open
              },
            }}
          />
        </LocaleProvider>,
      ),
    )
  },
})
