import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import { SplitDivider } from '../../apps/shell/src/renderer/src/SplitDivider'
import type { SplitViewState } from '../../apps/shell/src/shared/split-view-api'
import type { ViewportSize } from '../../apps/shell/src/renderer/src/split-view-geometry'

const root = createRoot(document.querySelector('#react-divider')!)
Object.assign(window, {
  aiOfficeTabs: { setSplitRatio: () => undefined },
  mountWorkspaceDivider(split: SplitViewState | null, viewport: ViewportSize) {
    flushSync(() =>
      root.render(
        <div className="app-frame">
          {split && <SplitDivider split={split} viewport={viewport} />}
        </div>,
      ),
    )
  },
})
