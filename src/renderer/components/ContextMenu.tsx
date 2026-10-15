import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { nextMenuIndex } from './menu-focus.js'

/** What every context menu shares: fitting on screen, the keyboard and closing. Callers only write the items. */

export interface MenuPosition {
  readonly x: number
  readonly y: number
}

const EDGE_MARGIN = 8
const SUBMENU_ATTRIBUTE = 'data-submenu'
const ENABLED_ITEM = '[role="menuitem"]:not(:disabled)'

function fitInWindow(position: MenuPosition, width: number, height: number): MenuPosition {
  return {
    x: Math.max(EDGE_MARGIN, Math.min(position.x, window.innerWidth - width - EDGE_MARGIN)),
    y: Math.max(EDGE_MARGIN, Math.min(position.y, window.innerHeight - height - EDGE_MARGIN)),
  }
}

function itemsOf(menu: Element): HTMLElement[] {
  return [...menu.querySelectorAll<HTMLElement>(ENABLED_ITEM)].filter(
    (item) => item.closest('[role="menu"]') === menu,
  )
}

function focusFirstItem(menu: Element | null): void {
  if (menu !== null) itemsOf(menu)[0]?.focus()
}

/** Arrows move within the menu that holds the focus, so a submenu keeps its own cycle. */
function moveFocus(event: React.KeyboardEvent): void {
  const target = event.target instanceof Element ? event.target : null
  const menu = target?.closest('[role="menu"]')
  if (menu === null || menu === undefined) return
  const items = itemsOf(menu)
  const next = nextMenuIndex(items.indexOf(target as HTMLElement), items.length, event.key)
  if (next === null) return
  event.preventDefault()
  items[next]?.focus()
}

function useDismiss(menu: React.RefObject<HTMLDivElement | null>, onClose: () => void): void {
  useEffect(() => {
    const dismiss = (event: Event): void => {
      if (event.target instanceof Node && menu.current?.contains(event.target) === true) return
      onClose()
    }
    // Escape inside an open submenu closes only the submenu, which handles it itself.
    const onKey = (event: KeyboardEvent): void => {
      const insideSubmenu =
        event.target instanceof Element && event.target.closest(`[${SUBMENU_ATTRIBUTE}]`) !== null
      if (event.key === 'Escape' && !insideSubmenu) onClose()
      // As in the system menus: Tab leaves the menu, and the focus goes back to where it was.
      if (event.key === 'Tab') {
        event.preventDefault()
        onClose()
      }
    }

    // `pointerdown`, not `click`, to close as soon as the pointer goes down elsewhere.
    document.addEventListener('pointerdown', dismiss, true)
    document.addEventListener('keydown', onKey, true)
    window.addEventListener('resize', onClose)
    return () => {
      document.removeEventListener('pointerdown', dismiss, true)
      document.removeEventListener('keydown', onKey, true)
      window.removeEventListener('resize', onClose)
    }
  }, [menu, onClose])
}

/**
 * Escape leaves the focus nowhere; give it back to where it was. Only while it is still on the page
 * body or in the menu: an action that moved it elsewhere keeps it.
 */
function useFocusReturn(menu: React.RefObject<HTMLDivElement | null>): void {
  useLayoutEffect(() => {
    const previous = document.activeElement
    focusFirstItem(menu.current)
    const element = menu.current
    return () => {
      const active = document.activeElement
      const lost = active === null || active === document.body || element?.contains(active) === true
      const returnable =
        previous instanceof HTMLElement && previous.isConnected && element?.contains(previous) !== true
      if (lost && returnable) previous.focus()
    }
  }, [menu])
}

export function ContextMenu({
  position,
  label,
  onClose,
  children,
}: {
  readonly position: MenuPosition
  /** Announced by the screen reader on open. */
  readonly label: string
  readonly onClose: () => void
  readonly children: React.ReactNode
}): React.JSX.Element {
  const menu = useRef<HTMLDivElement>(null)
  const [placement, setPlacement] = useState<MenuPosition>(position)

  // Measured only after drawing: the height depends on the system font.
  useLayoutEffect(() => {
    const element = menu.current
    if (element === null) return
    const { width, height } = element.getBoundingClientRect()
    setPlacement(fitInWindow(position, width, height))
  }, [position])

  useDismiss(menu, onClose)
  useFocusReturn(menu)

  return (
    <div
      ref={menu}
      className="ctx-menu"
      role="menu"
      aria-label={label}
      style={{ left: placement.x, top: placement.y }}
      onKeyDown={moveFocus}
    >
      {children}
    </div>
  )
}

/** Opens to the right, or to the left when the window has no room. */
function submenuPlacement(trigger: DOMRect, width: number, height: number): MenuPosition {
  const fitsRight = trigger.right + width + EDGE_MARGIN <= window.innerWidth
  return fitInWindow({ x: fitsRight ? trigger.right : trigger.left - width, y: trigger.top }, width, height)
}

export function ContextSubmenu({
  label,
  children,
}: {
  readonly label: string
  readonly children: React.ReactNode
}): React.JSX.Element {
  const trigger = useRef<HTMLButtonElement>(null)
  const submenu = useRef<HTMLDivElement>(null)
  const [open, setOpen] = useState(false)
  const [focusOnOpen, setFocusOnOpen] = useState(false)
  const [placement, setPlacement] = useState<MenuPosition | null>(null)

  useLayoutEffect(() => {
    const element = submenu.current
    if (!open || element === null || trigger.current === null) return
    const { width, height } = element.getBoundingClientRect()
    setPlacement(submenuPlacement(trigger.current.getBoundingClientRect(), width, height))
  }, [open])

  // Only once placed: a hidden element does not take the focus.
  useEffect(() => {
    if (open && focusOnOpen && placement !== null) focusFirstItem(submenu.current)
  }, [open, focusOnOpen, placement])

  const close = (): void => {
    setOpen(false)
    setFocusOnOpen(false)
    setPlacement(null)
  }
  /** The pointer passing by must not close a submenu the keyboard is working in. */
  const leave = (): void => {
    if (submenu.current?.contains(document.activeElement) !== true) close()
  }
  const openWithKeyboard = (): void => {
    setFocusOnOpen(true)
    setOpen(true)
  }
  const closeToTrigger = (event: React.KeyboardEvent): void => {
    if (event.key !== 'ArrowLeft' && event.key !== 'Escape') return
    event.preventDefault()
    event.stopPropagation()
    close()
    trigger.current?.focus()
  }

  return (
    <div className="ctx-menu__sub" onPointerEnter={() => setOpen(true)} onPointerLeave={leave}>
      <button
        ref={trigger}
        type="button"
        role="menuitem"
        aria-haspopup="menu"
        aria-expanded={open}
        className="ctx-menu__item ctx-menu__item--parent"
        onClick={openWithKeyboard}
        onKeyDown={(event) => {
          if (event.key !== 'ArrowRight') return
          event.preventDefault()
          openWithKeyboard()
        }}
      >
        {label}
        <span className="ctx-menu__arrow" aria-hidden="true">
          ▸
        </span>
      </button>
      {open && (
        <div
          ref={submenu}
          className="ctx-menu"
          role="menu"
          aria-label={label}
          data-submenu=""
          style={placement === null ? { visibility: 'hidden' } : { left: placement.x, top: placement.y }}
          onKeyDown={closeToTrigger}
        >
          {children}
        </div>
      )}
    </div>
  )
}

export function ContextMenuItem({
  children,
  onClick,
  disabled = false,
  strong = false,
}: {
  readonly children: React.ReactNode
  readonly onClick: () => void
  readonly disabled?: boolean
  /** Used by spellchecker suggestions. */
  readonly strong?: boolean
}): React.JSX.Element {
  return (
    <button
      type="button"
      role="menuitem"
      className={strong ? 'ctx-menu__item ctx-menu__item--suggestion' : 'ctx-menu__item'}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

export function ContextMenuSeparator(): React.JSX.Element {
  return <hr className="ctx-menu__sep" />
}
