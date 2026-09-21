import { useEffect, useLayoutEffect, useRef, useState } from 'react'

/**
 * A casca de um menu de contexto: onde ele aparece e quando ele fecha.
 *
 * Nasceu no menu da planilha e saiu de lá quando o editor de documentos ganhou o
 * seu — as duas coisas difíceis não têm nada de planilha nem de documento:
 *
 *  - **caber na tela**: a altura depende da fonte do sistema, então só medindo
 *    depois de desenhar se sabe se o menu cabe;
 *  - **fechar**: clique fora, `Escape` e redimensionamento da janela.
 *
 * Quem herda daqui só escreve os itens.
 */

export interface MenuPosition {
  readonly x: number
  readonly y: number
}

/** Distância mínima da borda da janela, para o menu não encostar. */
const EDGE_MARGIN = 8

export function ContextMenu({
  position,
  label,
  onClose,
  children,
}: {
  readonly position: MenuPosition
  /** Nome acessível do menu: é o que o leitor de tela anuncia ao abrir. */
  readonly label: string
  readonly onClose: () => void
  readonly children: React.ReactNode
}): React.JSX.Element {
  const menu = useRef<HTMLDivElement>(null)
  const [placement, setPlacement] = useState<MenuPosition>(position)

  // Medir depois de desenhar é o único jeito de saber se o menu cabe: a altura
  // depende da fonte do sistema, não de constante nossa.
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

    // `pointerdown` e não `click`: fechar só no clique deixaria o menu aberto
    // enquanto o botão está pressionado em outro lugar.
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
  /** Destaque de "é este que você quer" — as sugestões do corretor o usam. */
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
