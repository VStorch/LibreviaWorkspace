import { useCallback, useMemo, useRef, useState } from 'react'
import { RevoGrid, type ColumnRegular } from '@revolist/react-datagrid'
import type {
  AfterEditEvent,
  BeforeSaveDataDetails,
  ChangedRange,
  FocusAfterRenderEvent,
  RevoGridCustomEvent,
} from '@revolist/revogrid'
import { formatCell } from '@services/spreadsheet/format.js'
import { DEFAULT_COLUMN_WIDTH, columnName, getCell, type Sheet } from '@services/spreadsheet/model.js'
import {
  normalizeRange,
  rangeContains,
  singleCell,
  toggleStyle,
  writeText,
  type Range,
} from '@services/spreadsheet/edit.js'
import { fillRange } from '@services/spreadsheet/fill.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'
import { FormulaBar } from './FormulaBar.js'
import { SpreadsheetToolbar } from './SpreadsheetToolbar.js'
import { SheetContextMenu, type MenuPosition } from './SheetContextMenu.js'
import { cellStyleOf } from './cell-style.js'
import { gridPositionOf } from './grid-position.js'
import { useFormatShortcuts } from './useFormatShortcuts.js'
import { useTypeAhead } from './useTypeAhead.js'

/**
 * Editor de planilhas.
 *
 * O grid é o `@revolist/revogrid` (MIT), escolhido no spike da §4.2 do plano:
 10 mil linhas em 67 ms, com 338 células no DOM. **Mesclagem de células ficou
 * fora desta fase** — toda biblioteca madura a cobra, e a decisão está
 * registrada no plano.
 *
 * O grid trabalha com linhas de objeto; o nosso modelo é um **mapa esparso** por
 * referência A1. A tradução mora aqui, e só aqui: uma planilha de 10 mil linhas
 * com trinta células preenchidas continua ocupando trinta entradas na memória e
 * no arquivo.
 */

/** Uma linha como o grid espera: colunas indexadas por `c0`, `c1`… */
type GridRow = Record<string, string>

export function SpreadsheetEditor({
  sheet,
  onChange,
  onStructure,
  readOnly = false,
}: {
  sheet: Sheet
  onChange: (sheet: Sheet) => void
  /**
   * Inserir e excluir linha ou coluna é operação da **pasta**, não da planilha:
   * uma linha inserida aqui muda o significado de `=Dados!A5` escrita em outra
   * aba. Por isso ela sobe até o estado em vez de virar uma nova `Sheet`.
   */
  onStructure: (change: StructuralChange) => void
  /** Trava a edição sem esconder nada: ler e rolar continuam funcionando. */
  readOnly?: boolean
}): React.JSX.Element {
  // O modelo mais recente, para o handler de edição não capturar um estado
  // velho entre renderizações.
  const current = useRef(sheet)
  current.current = sheet

  // A seleção mora aqui, e não no grid, porque é a barra de ferramentas que
  // precisa dela — e ela sobrevive à troca de foco entre grade e botões.
  const [range, setRange] = useState<Range>(() => singleCell(0, 0))
  const selection = useRef(range)
  selection.current = range

  const [menu, setMenu] = useState<MenuPosition | null>(null)

  const typeAhead = useTypeAhead(readOnly)

  /**
   * Portão único do somente leitura.
   *
   * A grade tem `readonly`, mas ela não é a única porta: barra de fórmulas,
   * barra de ferramentas, menu de contexto e os atalhos de formatação também
   * escrevem. Passar tudo por aqui é o que garante que nenhuma delas escape —
   * uma trava que depende de lembrar de sete lugares não é uma trava.
   */
  const applyChange = useCallback(
    (next: Sheet) => {
      if (!readOnly) onChange(next)
    },
    [onChange, readOnly],
  )

  const applyStructure = useCallback(
    (change: StructuralChange) => {
      if (!readOnly) onStructure(change)
    },
    [onStructure, readOnly],
  )

  const columns = useMemo<ColumnRegular[]>(
    () =>
      Array.from({ length: sheet.columnCount }, (_, index) => {
        const column: ColumnRegular = {
          prop: `c${index}`,
          name: columnName(index),
          size: sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH,
          resizable: true,
          // A aparência é aplicada por célula, na hora de desenhar. O estilo
          // vive no modelo, não no DOM: é o que faz a formatação sobreviver ao
          // salvamento e à rolagem, que descarta e recria as células.
          cellProperties: ({ rowIndex }) => ({
            style: cellStyleOf(current.current, rowIndex, index),
          }),
        }
        if (index < sheet.frozenColumns) column.pin = 'colPinStart'
        return column
      }),
    // `sheet` inteiro: mudar formatação precisa redesenhar, e o estilo é lido
    // de `current` no momento do desenho.
    [sheet],
  )

  /**
   * O grid precisa de uma linha por posição visível, mesmo vazia — é o que dá a
   * grade. As linhas são geradas sob demanda a partir do mapa esparso, então a
   * memória continua proporcional ao conteúdo, não à área.
   */
  const source = useMemo<GridRow[]>(() => {
    const rows: GridRow[] = []
    for (let row = 0; row < sheet.rowCount; row++) {
      const cells: GridRow = {}
      for (let column = 0; column < sheet.columnCount; column++) {
        cells[`c${column}`] = formatCell(getCell(sheet, row, column))
      }
      rows.push(cells)
    }
    return rows
  }, [sheet])

  /** Escreve pelo nome da coluna que o grid usa (`c0`, `c1`…). */
  const writeAt = useCallback((sheet: Sheet, row: number, prop: string, raw: unknown): Sheet => {
    const column = Number.parseInt(prop.slice(1), 10)
    if (!Number.isInteger(column)) return sheet

    return writeText(sheet, row, column, typeof raw === 'string' ? raw : String(raw ?? ''))
  }, [])

  /**
   * O evento cobre **dois** casos: uma célula editada e um intervalo colado. O
   * segundo é o que faz colar de outra planilha funcionar, e ele chega com
   * forma diferente — tratar só o primeiro descartaria a colagem em silêncio.
   */
  const handleEdit = useCallback(
    (event: RevoGridCustomEvent<AfterEditEvent>) => {
      const detail = event.detail
      let updated = current.current

      // `newRange` só existe na colagem de intervalo. `data` não serve como
      // discriminante: os dois tipos a declaram.
      if ('newRange' in detail) {
        for (const [row, values] of Object.entries(detail.data)) {
          for (const [prop, value] of Object.entries(values as Record<string, unknown>)) {
            updated = writeAt(updated, Number(row), prop, value)
          }
        }
      } else {
        updated = writeAt(updated, detail.rowIndex, String(detail.prop), detail.val)
      }

      applyChange(updated)

      // Colar e confirmar com o mouse não passam pelo Enter: a janela também
      // abre aqui, e fecha quando o foco chega na célula seguinte.
      typeAhead.begin()
    },
    [applyChange, typeAhead, writeAt],
  )

  /**
   * Abrir a célula para edição mostra a **fórmula**, e não o resultado.
   *
   * A grade exibe o valor calculado, que é o certo para ler. Mas entrar na
   * célula com F2 e sair sem querer gravaria esse número por cima da fórmula —
   * uma perda silenciosa a cada toque acidental. O grid usa `detail.val` como
   * conteúdo inicial do editor, então é ele que precisa ser trocado.
   */
  const handleEditStart = useCallback((event: RevoGridCustomEvent<BeforeSaveDataDetails>) => {
    const detail = event.detail
    const column = Number.parseInt(String(detail.prop).slice(1), 10)
    if (!Number.isInteger(column)) return

    const formula = getCell(current.current, detail.rowIndex, column)?.formula
    if (formula !== undefined) detail.val = formula
  }, [])

  const handleResize = useCallback(
    (event: CustomEvent<Record<number, ColumnRegular>>) => {
      const widths = { ...current.current.columnWidths }
      for (const [index, column] of Object.entries(event.detail)) {
        if (column.size !== undefined) widths[Number(index)] = column.size
      }
      applyChange({ ...current.current, columnWidths: widths })
    },
    [applyChange],
  )

  const handleFocus = useCallback(
    (event: RevoGridCustomEvent<FocusAfterRenderEvent>) => {
      const { rowIndex, colIndex } = event.detail
      if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex)) return

      setRange(singleCell(rowIndex, colIndex))
      typeAhead.settle()
    },
    [typeAhead],
  )

  const handleRange = useCallback((event: RevoGridCustomEvent<ChangedRange>) => {
    const area = event.detail.newRange
    if (area === null) return
    setRange(normalizeRange({ fromRow: area.y, fromColumn: area.x, toRow: area.y1, toColumn: area.x1 }))
  }, [])

  /**
   * Alça de preenchimento.
   *
   * O grid preencheria copiando o **texto exibido**, que para uma fórmula é o
   * resultado já calculado — arrastar `=B2*C2` para baixo repetiria o número da
   * primeira linha, e o erro só apareceria no fechamento do mês. Por isso o
   * comportamento padrão é cancelado e o preenchimento é feito sobre o modelo,
   * onde a fórmula existe e pode ser deslocada.
   */
  const handleAutofill = useCallback(
    (event: RevoGridCustomEvent<ChangedRange>) => {
      const { oldRange, newRange } = event.detail
      if (oldRange === null || newRange === null) return

      event.preventDefault()
      applyChange(
        fillRange(
          current.current,
          { fromRow: oldRange.y, fromColumn: oldRange.x, toRow: oldRange.y1, toColumn: oldRange.x1 },
          { fromRow: newRange.y, fromColumn: newRange.x, toRow: newRange.y1, toColumn: newRange.x1 },
        ),
      )
    },
    [applyChange],
  )

  /**
   * Botão direito: o menu age sobre a seleção quando o clique cai **dentro**
   * dela, e sobre a célula clicada quando cai fora — a regra do Excel. Clicar
   * fora e mesmo assim operar sobre a seleção antiga excluiria a linha errada.
   */
  const handleContextMenu = useCallback((event: React.MouseEvent<HTMLDivElement>) => {
    event.preventDefault()

    // Cabeçalho de coluna e área vazia não trazem posição: aí vale a seleção.
    const position = gridPositionOf(event.nativeEvent)
    if (position !== null && !rangeContains(selection.current, position.row, position.column)) {
      setRange(singleCell(position.row, position.column))
    }

    setMenu({ x: event.clientX, y: event.clientY })
  }, [])

  const closeMenu = useCallback(() => setMenu(null), [])

  useFormatShortcuts((key) => applyChange(toggleStyle(current.current, selection.current, key)))

  return (
    <div className="sheet" onContextMenu={handleContextMenu}>
      <SpreadsheetToolbar sheet={sheet} range={range} onChange={applyChange} />

      <FormulaBar
        sheet={sheet}
        range={range}
        onCommit={(text) => applyChange(writeText(current.current, range.fromRow, range.fromColumn, text))}
      />

      {menu !== null && (
        <SheetContextMenu
          sheet={sheet}
          range={range}
          position={menu}
          onChange={onChange}
          onStructure={applyStructure}
          onClose={closeMenu}
        />
      )}

      <RevoGrid
        columns={columns}
        source={source}
        theme="compact"
        resize={true}
        range={true}
        readonly={readOnly}
        rowHeaders={true}
        useClipboard={true}
        rowDefinitions={Array.from({ length: sheet.frozenRows }, (_, index) => ({
          type: 'rowPinStart' as const,
          index,
          size: 24,
        }))}
        onAfteredit={handleEdit}
        onBeforeeditstart={handleEditStart}
        onAftercolumnresize={handleResize}
        onAfterfocus={handleFocus}
        onBeforerange={handleRange}
        onBeforeautofill={handleAutofill}
      />
    </div>
  )
}
