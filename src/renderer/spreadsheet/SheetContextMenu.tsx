import type { Sheet } from '@services/spreadsheet/model.js'
import { clearContents, type Range } from '@services/spreadsheet/edit.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'
import {
  ContextMenu,
  ContextMenuItem,
  ContextMenuSeparator,
  type MenuPosition,
} from '../components/ContextMenu.js'

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
 *
 * Onde o menu aparece e quando ele fecha mora em `components/ContextMenu.tsx`:
 * aquilo não tem nada de planilha, e o editor de documentos usa o mesmo.
 */

export type { MenuPosition }

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
    <ContextMenu position={position} label="Ações da planilha" onClose={onClose}>
      <ContextMenuItem onClick={structural({ kind: 'insertRows', at: range.fromRow, count: rows })}>
        Inserir {plural(rows, 'linha', 'linhas')} acima
      </ContextMenuItem>
      <ContextMenuItem onClick={structural({ kind: 'insertRows', at: range.toRow + 1, count: rows })}>
        Inserir {plural(rows, 'linha', 'linhas')} abaixo
      </ContextMenuItem>
      <ContextMenuItem onClick={structural({ kind: 'deleteRows', at: range.fromRow, count: rows })}>
        Excluir {plural(rows, 'linha', 'linhas')}
      </ContextMenuItem>

      <ContextMenuSeparator />

      <ContextMenuItem onClick={structural({ kind: 'insertColumns', at: range.fromColumn, count: columns })}>
        Inserir {plural(columns, 'coluna', 'colunas')} à esquerda
      </ContextMenuItem>
      <ContextMenuItem
        onClick={structural({ kind: 'insertColumns', at: range.toColumn + 1, count: columns })}
      >
        Inserir {plural(columns, 'coluna', 'colunas')} à direita
      </ContextMenuItem>
      <ContextMenuItem onClick={structural({ kind: 'deleteColumns', at: range.fromColumn, count: columns })}>
        Excluir {plural(columns, 'coluna', 'colunas')}
      </ContextMenuItem>

      <ContextMenuSeparator />

      <ContextMenuItem onClick={run((s) => clearContents(s, range))}>Limpar conteúdo</ContextMenuItem>
    </ContextMenu>
  )
}

/** "1 linha", "3 linhas" — o número aparece só quando não é um. */
function plural(count: number, one: string, many: string): string {
  return count === 1 ? one : `${count} ${many}`
}
