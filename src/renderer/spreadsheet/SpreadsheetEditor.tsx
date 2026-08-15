import { useCallback, useMemo, useRef } from 'react'
import { RevoGrid, type ColumnRegular } from '@revolist/react-datagrid'
import type { AfterEditEvent, RevoGridCustomEvent } from '@revolist/revogrid'
import { formatCell, parseInput } from '@services/spreadsheet/format.js'
import {
  DEFAULT_COLUMN_WIDTH,
  cellRef,
  columnName,
  getCell,
  setCell,
  type Cell,
  type Sheet,
} from '@services/spreadsheet/model.js'

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
}: {
  sheet: Sheet
  onChange: (sheet: Sheet) => void
}): React.JSX.Element {
  // O modelo mais recente, para o handler de edição não capturar um estado
  // velho entre renderizações.
  const current = useRef(sheet)
  current.current = sheet

  const columns = useMemo<ColumnRegular[]>(
    () =>
      Array.from({ length: sheet.columnCount }, (_, index) => {
        const column: ColumnRegular = {
          prop: `c${index}`,
          name: columnName(index),
          size: sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH,
          resizable: true,
        }
        if (index < sheet.frozenColumns) column.pin = 'colPinStart'
        return column
      }),
    [sheet.columnCount, sheet.columnWidths, sheet.frozenColumns],
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

  /** Aplica um texto digitado ou colado a uma posição. */
  const write = useCallback((sheet: Sheet, row: number, prop: string, raw: unknown): Sheet => {
    const column = Number.parseInt(prop.slice(1), 10)
    if (!Number.isInteger(column)) return sheet

    const previous = getCell(sheet, row, column)
    const parsed = parseInput(typeof raw === 'string' ? raw : String(raw ?? ''))

    // O formato reconhecido na digitação não apaga o que o usuário escolheu à
    // mão: quem já pintou a célula de moeda não quer perder isso ao redigitar.
    const cell: Cell = { value: parsed.value }
    const style = previous?.style ?? parsed.style
    // Atribuição condicional por causa de `exactOptionalPropertyTypes`: a
    // propriedade ausente não é o mesmo que a propriedade indefinida.
    const withStyle: Cell = style === undefined ? cell : { ...cell, style }

    return setCell(sheet, row, column, withStyle)
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
            updated = write(updated, Number(row), prop, value)
          }
        }
      } else {
        updated = write(updated, detail.rowIndex, String(detail.prop), detail.val)
      }

      onChange(updated)
    },
    [onChange, write],
  )

  const handleResize = useCallback(
    (event: CustomEvent<Record<number, ColumnRegular>>) => {
      const widths = { ...current.current.columnWidths }
      for (const [index, column] of Object.entries(event.detail)) {
        if (column.size !== undefined) widths[Number(index)] = column.size
      }
      onChange({ ...current.current, columnWidths: widths })
    },
    [onChange],
  )

  return (
    <div className="sheet">
      <RevoGrid
        columns={columns}
        source={source}
        theme="compact"
        resize={true}
        range={true}
        rowHeaders={true}
        useClipboard={true}
        rowDefinitions={Array.from({ length: sheet.frozenRows }, (_, index) => ({
          type: 'rowPinStart' as const,
          index,
          size: 24,
        }))}
        onAfteredit={handleEdit}
        onAftercolumnresize={handleResize}
      />

      <div className="sheet__ref" aria-live="polite">
        {cellRef(0, 0)}
      </div>
    </div>
  )
}
