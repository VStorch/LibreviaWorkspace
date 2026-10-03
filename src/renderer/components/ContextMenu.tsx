import { useEffect, useLayoutEffect, useRef, useState } from 'react'

/** O que é comum a todo menu de contexto: caber na tela e fechar. Quem herda só escreve os itens. */

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
  /** O que o leitor de tela anuncia ao abrir. */
  readonly label: string
  readonly onClose: () => void
  readonly children: React.ReactNode
}): React.JSX.Element {
  const menu = useRef<HTMLDivElement>(null)
  const [placement, setPlacement] = useState<MenuPosition>(position)

  // Só medindo depois de desenhar: a altura depende da fonte do sistema.
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

    // `pointerdown`, e não `click`, para fechar já ao pressionar em outro lugar.
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
  /** As sugestões do corretor o usam. */
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
