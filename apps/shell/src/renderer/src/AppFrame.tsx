import { useEffect, useState } from 'react'
import type { SplitViewState } from '../../shared/split-view-api'
import { useFirstConnectionNotice } from './FirstConnectionNotice'
import { Home } from './Home'
import { Onboarding } from './Onboarding'
import type { ViewportSize } from './split-view-geometry'
import { SplitDivider } from './SplitDivider'
import { TabBar } from './TabBar'
import { useSubscriptionConnection } from './subscription-client'
import { workspaceUiScale } from '../../shared/workspace-layout'

interface AppFrameProps {
  /** resolved before first paint (main.tsx) so home never flashes under the overlay */
  initialOnboardingSeen: boolean
}

/** The shell renderer's viewport: the same CSS px space the pane bounds use. */
function useViewportSize(): ViewportSize {
  const readSize = (): ViewportSize => ({
    width: window.innerWidth,
    height: window.innerHeight,
    uiScale: workspaceUiScale(
      Number.parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--ui-scale')),
    ),
  })
  const [size, setSize] = useState<ViewportSize>(readSize)
  useEffect(() => {
    const onResize = (): void => setSize(readSize())
    window.addEventListener('resize', onResize)
    const observer = new MutationObserver(onResize)
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['style'] })
    return () => {
      window.removeEventListener('resize', onResize)
      observer.disconnect()
    }
  }, [])
  return size
}

/** Read the split once, then follow the main process' change notices. */
function useSplitView(): SplitViewState | null {
  const [split, setSplit] = useState<SplitViewState | null>(null)
  useEffect(() => {
    let alive = true
    // until the split wiring is registered a bare read rejects: "not split"
    void window.aiOfficeTabs.getSplit().then(
      (state) => {
        if (alive) setSplit(state)
      },
      () => undefined,
    )
    const off = window.aiOfficeTabs.onSplitChanged(setSplit)
    return () => {
      alive = false
      off()
    }
  }, [])
  return split
}

export function AppFrame({ initialOnboardingSeen }: AppFrameProps) {
  const [homeActive, setHomeActive] = useState(true)
  const [showOnboarding, setShowOnboarding] = useState(!initialOnboardingSeen)
  const split = useSplitView()
  const viewport = useViewportSize()
  // one live view of the subscription connection for the home screen, the
  // onboarding and the first-connection notice
  const connection = useSubscriptionConnection()
  const notice = useFirstConnectionNotice(connection.statuses)

  useEffect(() => {
    const applyTabs = (tabs: Awaited<ReturnType<typeof window.aiOfficeTabs.list>>) => {
      const active = tabs.find((tab) => tab.active)
      setHomeActive(!active || active.kind === 'home')
    }
    void window.aiOfficeTabs.list().then(applyTabs)
    return window.aiOfficeTabs.onChanged(applyTabs)
  }, [])

  const finishOnboarding = async (): Promise<boolean> => {
    try {
      const persisted = await window.aiOffice.setOnboardingSeen()
      if (!persisted) return false
      setShowOnboarding(false)
      return true
    } catch {
      return false
    }
  }

  return (
    <div className="app-frame">
      <TabBar split={split} />
      {/* divider sits in the gap between the two pane views — the one slice of
       * the content area they do not cover. Home paints this document instead
       * (the panes are only hidden there, not closed), so no seam over it. */}
      {split && !homeActive && <SplitDivider split={split} viewport={viewport} />}
      {/* docs/sheets tabs render as WebContentsView children of this window, positioned
       * by the main process to cover this area — only Home paints its own content here. */}
      <div className="app-frame-content" style={{ visibility: homeActive ? 'visible' : 'hidden' }}>
        <Home
          connection={connection}
          aiSettings={notice.settings}
          // during the onboarding its connect step shows the notice instead
          noticeDue={notice.due && !showOnboarding}
          onNoticeAcknowledge={notice.acknowledge}
        />
      </div>
      {/* editor WebContentsViews paint above ALL shell DOM, so the overlay only
       * renders while the home tab is active — it comes back when home does */}
      {showOnboarding && homeActive && (
        <Onboarding
          onDone={finishOnboarding}
          connection={connection}
          noticeDue={notice.due}
          onNoticeAcknowledge={notice.acknowledge}
        />
      )}
    </div>
  )
}
