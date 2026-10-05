import { useEffect, useLayoutEffect, useRef, useState } from 'react'

/** What every context menu shares: fitting on screen and closing. Callers only write the items. */

export interface MenuPosition {
  readonly x: number
  readonly y: number
}

const EDGE_MARGIN = 8

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
    setPlacement({
      x: Math.max(EDGE_MARGIN, Math.min(position.x, window.innerWidth - width - EDGE_MARGIN)),
      y: Math.max(EDGE_MARGIN, Math.min(position.y, window.innerHeight - height - EDGE_MARGIN)),
    })
  }, [position])

  useEffect(() => {
    const dismiss = (event: Event): void => {
      if (event.target instanceof Node && menu.current?.contains(event.target) === true) return
      onClose()
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') onClose()
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
  }, [onClose])

  return (
    <div
      ref={menu}
      className="ctx-menu"
      role="menu"
      aria-label={label}
      style={{ left: placement.x, top: placement.y }}
    >
      {children}
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
