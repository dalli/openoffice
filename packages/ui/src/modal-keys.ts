/**
 * Shared modal keyboard behavior for the renderer dialogs: Esc closes (stopped
 * so it never reaches global listeners), Tab cycles inside the modal, and the
 * first form control gets focus on mount unless something inside is already
 * focused. Spread the returned ref/onKeyDown onto the modal backdrop element.
 */
import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

const FOCUSABLE =
  'a[href], button, input, textarea, select, [contenteditable]:not([contenteditable="false"]), [tabindex]'

/** Editing and native form controls own Enter and arrow keys, including IME input. */
export function isEditingKeyTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false
  if (target.closest('input, textarea, select')) return true
  const editable = target.closest('[contenteditable]')
  return editable !== null && editable.getAttribute('contenteditable') !== 'false'
}

function available(el: HTMLElement): boolean {
  if (el.matches(':disabled') || el.closest('[inert], [hidden], [aria-hidden="true"]')) return false
  for (let node: HTMLElement | null = el; node; node = node.parentElement) {
    const style = getComputedStyle(node)
    if (style.display === 'none' || style.visibility === 'hidden') return false
  }
  return true
}

function focusableItems(container: HTMLElement): HTMLElement[] {
  return [...container.querySelectorAll<HTMLElement>(FOCUSABLE)].filter(
    (el) =>
      available(el) &&
      (el.tabIndex >= 0 || (el.hasAttribute('contenteditable') && !el.hasAttribute('tabindex'))),
  )
}

// Child effects can mount before their parent. Prefer the innermost modal,
// then the latest sibling, so a notice above another dialog owns the keyboard.
const modals: HTMLElement[] = []
function currentModal(): HTMLElement | undefined {
  for (let i = modals.length - 1; i >= 0; i--) {
    const el = modals[i]!
    if (
      el.isConnected &&
      !modals.some((other) => other !== el && other.isConnected && el.contains(other))
    )
      return el
  }
  return undefined
}

export function useModalKeys(
  onClose: () => void,
  options?: { restoreFocus?: boolean; active?: boolean; initialFocus?: 'first' | 'container' },
) {
  const ref = useRef<HTMLDivElement>(null)
  const restoreFocus = options?.restoreFocus ?? false
  const active = options?.active ?? true
  const initialFocus = options?.initialFocus ?? 'first'

  useEffect(() => {
    const el = ref.current
    const previous =
      restoreFocus && document.activeElement instanceof HTMLElement ? document.activeElement : null
    return () => {
      const index = el ? modals.indexOf(el) : -1
      if (index >= 0) modals.splice(index, 1)
      if (previous && previous.isConnected) previous.focus()
    }
  }, [restoreFocus])

  useEffect(() => {
    const el = ref.current
    if (!el || !active) return
    modals.push(el)
    const focusInside = () => {
      const first = initialFocus === 'container' ? undefined : focusableItems(el)[0]
      ;(first ?? el).focus({ preventScroll: true })
    }
    const onFocus = (event: FocusEvent) => {
      if (currentModal() === el && !el.contains(event.target as Node)) focusInside()
    }
    if (currentModal() === el && !el.contains(document.activeElement)) focusInside()
    document.addEventListener('focusin', onFocus)
    return () => {
      document.removeEventListener('focusin', onFocus)
      const index = modals.indexOf(el)
      if (index >= 0) modals.splice(index, 1)
    }
  }, [active, initialFocus])

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (!active || e.defaultPrevented || currentModal() !== ref.current) return
    if (e.nativeEvent.isComposing || e.nativeEvent.keyCode === 229) return
    if (e.key === 'Escape') {
      e.preventDefault()
      e.stopPropagation()
      onClose()
      return
    }
    if (e.key === 'Tab') trapTab(ref.current, e)
  }

  return { ref, onKeyDown }
}

/** Cycle Tab/Shift+Tab inside a container (pure DOM helper, exported for tests). */
export function trapTab(
  container: HTMLElement | null,
  e: { key: string; shiftKey: boolean; preventDefault: () => void },
): void {
  if (!container || e.key !== 'Tab') return
  const items = focusableItems(container)
  if (items.length === 0) {
    e.preventDefault()
    return
  }
  const first = items[0]!
  const last = items[items.length - 1]!
  const active = document.activeElement as HTMLElement | null
  // A backdrop that holds focus itself is inside the trap, so name it alongside
  // the first/last control: without it Shift+Tab from the backdrop is left to
  // the browser and walks out of the modal.
  if (e.shiftKey && (active === first || !items.includes(active!))) {
    e.preventDefault()
    last.focus()
  } else if (!e.shiftKey && (active === last || !items.includes(active!))) {
    e.preventDefault()
    first.focus()
  }
}
