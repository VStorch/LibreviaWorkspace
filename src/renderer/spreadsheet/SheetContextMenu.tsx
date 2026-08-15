import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Sheet } from '@services/spreadsheet/model.js'
import { clearContents, type Range } from '@services/spreadsheet/edit.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'

/**
 * Menu de contexto da planilha.
 *
 * Inserir e excluir linha ou coluna não cabe na barra de ferramentas: são
 * operações sobre a **posição** clicada, e o gesto que todo mundo já conhece é
 * o botão direito.
 *
 * A quantidade vem da seleção, como no Excel e no Google Sheets: com três
 * linhas selecionadas, "inserir acima" insere três. Inserir uma de cada vez
 * seria o comportamento de um editor que não sabe o que está selecionado.
 */

export interface MenuPosition {
  readonly x: number
  readonly y: number
}

/** Distância mínima da borda da janela, para o menu não encostar. */
const EDGE_MARGIN = 8

export function SheetContextMenu({
  sheet,
  range,
  position,
  onChange,
  onStructure,
  onClose,
}: {
  sheet: Sheet
  range: Range
  position: MenuPosition
  onChange: (sheet: Sheet) => void
  onStructure: (change: StructuralChange) => void
  onClose: () => void
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

  const rows = range.toRow - range.fromRow + 1
  const columns = range.toColumn - range.fromColumn + 1

  const run = (operation: (sheet: Sheet) => Sheet) => () => {
    onChange(operation(sheet))
    onClose()
  }

  const structural = (change: StructuralChange) => () => {
    onStructure(change)
    onClose()
  }

  return (
    <div
      ref={menu}
      className="sheet-menu"
      role="menu"
      aria-label="Ações da planilha"
      style={{ left: placement.x, top: placement.y }}
    >
      <Item onClick={structural({ kind: 'insertRows', at: range.fromRow, count: rows })}>
        Inserir {plural(rows, 'linha', 'linhas')} acima
      </Item>
      <Item onClick={structural({ kind: 'insertRows', at: range.toRow + 1, count: rows })}>
        Inserir {plural(rows, 'linha', 'linhas')} abaixo
      </Item>
      <Item onClick={structural({ kind: 'deleteRows', at: range.fromRow, count: rows })}>
        Excluir {plural(rows, 'linha', 'linhas')}
      </Item>

      <hr className="sheet-menu__sep" />

      <Item onClick={structural({ kind: 'insertColumns', at: range.fromColumn, count: columns })}>
        Inserir {plural(columns, 'coluna', 'colunas')} à esquerda
      </Item>
      <Item onClick={structural({ kind: 'insertColumns', at: range.toColumn + 1, count: columns })}>
        Inserir {plural(columns, 'coluna', 'colunas')} à direita
      </Item>
      <Item onClick={structural({ kind: 'deleteColumns', at: range.fromColumn, count: columns })}>
        Excluir {plural(columns, 'coluna', 'colunas')}
      </Item>

      <hr className="sheet-menu__sep" />

      <Item onClick={run((s) => clearContents(s, range))}>Limpar conteúdo</Item>
    </div>
  )
}

/** "1 linha", "3 linhas" — o número aparece só quando não é um. */
function plural(count: number, one: string, many: string): string {
  return count === 1 ? one : `${count} ${many}`
}

function Item({ children, onClick }: { children: React.ReactNode; onClick: () => void }): React.JSX.Element {
  return (
    <button type="button" role="menuitem" className="sheet-menu__item" onClick={onClick}>
      {children}
    </button>
  )
}
