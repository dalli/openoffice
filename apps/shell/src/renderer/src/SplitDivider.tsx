import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { PointerEvent as ReactPointerEvent } from 'react'
import type { SplitViewState } from '../../shared/split-view-api'
import {
  SPLIT_EVEN_RATIO,
  dividerRect,
  ratioForDragOffset,
  splitRatioLimits,
  type ViewportSize,
} from './split-view-geometry'

interface DividerDrag {
  pointerId: number
  /** pointer position minus the divider's leading edge at pointer-down */
  grabOffset: number
  frame: number | null
  pending: number | null
}

/**
 * The two panes are WebContentsViews that paint above shell DOM, so the only
 * chrome the divider can live in is the gap the main process leaves between
 * them. Its rectangle is recomputed from the viewport and the split ratio.
 */
export function SplitDivider({
  split,
  viewport,
}: {
  split: SplitViewState
  viewport: ViewportSize
}) {
  const dragRef = useRef<DividerDrag | null>(null)
  // during a drag the divider follows the pointer here so it never lags the
  // main process re-laying the views out; the answered state takes over after
  const [dragRatio, setDragRatio] = useState<number | null>(null)
  const { min, max } = splitRatioLimits(viewport, split.orientation)
  const ratio = Math.min(max, Math.max(min, dragRatio ?? split.ratio))
  const rect = dividerRect(viewport, split.orientation, ratio)

  // a queued frame must never fire into an unmounted divider (e.g. the split
  // closed mid-drag)
  useEffect(() => {
    return () => {
      const drag = dragRef.current
      if (drag?.frame != null) cancelAnimationFrame(drag.frame)
      dragRef.current = null
    }
  }, [])

  const pointerAlongAxis = (event: ReactPointerEvent<HTMLDivElement>): number =>
    split.orientation === 'side' ? event.clientX : event.clientY

  const send = (value: number): void => window.aiOfficeTabs.setSplitRatio(value)

  const endDrag = (): void => {
    const drag = dragRef.current
    if (drag?.frame != null) cancelAnimationFrame(drag.frame)
    dragRef.current = null
    setDragRatio(null)
  }

  const handlePointerDown = (event: ReactPointerEvent<HTMLDivElement>): void => {
    if (event.button !== 0) return
    const current = dividerRect(viewport, split.orientation, ratio)
    const leading = split.orientation === 'side' ? current.left : current.top
    dragRef.current = {
      pointerId: event.pointerId,
      grabOffset: pointerAlongAxis(event) - leading,
      frame: null,
      pending: null,
    }
    event.currentTarget.setPointerCapture(event.pointerId)
    setDragRatio(split.ratio)
  }

  const handlePointerMove = (event: ReactPointerEvent<HTMLDivElement>): void => {
    const drag = dragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    const target = pointerAlongAxis(event) - drag.grabOffset
    const next = ratioForDragOffset(target, viewport, split.orientation)
    // past a pane's minimum the drag means "close that side": say so once and
    // let the split notice remove the divider
    if (next === 0 || next === 1) {
      endDrag()
      send(next)
      return
    }
    setDragRatio(next)
    drag.pending = next
    // every accepted ratio re-lays both views in the main process, so the
    // pointer stream is sampled to one send per animation frame
    if (drag.frame == null) {
      drag.frame = requestAnimationFrame(() => {
        const current = dragRef.current
        if (!current) return
        current.frame = null
        const pending = current.pending
        current.pending = null
        if (pending != null) send(pending)
      })
    }
  }

  const handlePointerUp = (event: ReactPointerEvent<HTMLDivElement>): void => {
    const drag = dragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    const target = pointerAlongAxis(event) - drag.grabOffset
    // the release position supersedes the last throttled sample
    const final = ratioForDragOffset(target, viewport, split.orientation)
    endDrag()
    send(final)
  }

  return createPortal(
    <div
      className={`split-divider ${split.orientation === 'side' ? 'side' : 'stack'}`}
      style={{ left: rect.left, top: rect.top, width: rect.width, height: rect.height }}
      role="separator"
      aria-orientation={split.orientation === 'side' ? 'vertical' : 'horizontal'}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onPointerCancel={endDrag}
      onLostPointerCapture={endDrag}
      onDoubleClick={() => send(SPLIT_EVEN_RATIO)}
    />,
    document.body,
  )
}
