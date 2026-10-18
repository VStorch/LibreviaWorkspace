import { useMemo } from 'react'
import { pasteText } from '@services/spreadsheet/clipboard.js'
import { cellsIn, describeRange, editableText, type Range } from '@services/spreadsheet/edit.js'
import { quoteSheet } from '@services/spreadsheet/formula/references.js'
import { columnName, getCell, type Sheet } from '@services/spreadsheet/model.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { discardSheetProposal, useSheetMcpProposals, type PendingSheetProposal } from './mcp-sheet.js'

/** Enough to see what goes where; the whole proposal goes in on apply. */
const PREVIEW_ROWS = 5
const PREVIEW_COLUMNS = 6

function targetOf({ row, column, rows }: PendingSheetProposal): Range {
  const width = Math.max(...rows.map((cells) => cells.length))
  return { fromRow: row, fromColumn: column, toRow: row + rows.length - 1, toColumn: column + width - 1 }
}

function filledIn(sheet: Sheet, target: Range): number {
  let count = 0
  for (const { row, column } of cellsIn(target)) {
    if (editableText(getCell(sheet, row, column)) !== '') count++
  }
  return count
}

function Preview({ proposal }: { readonly proposal: PendingSheetProposal }): React.JSX.Element {
  const t = useT()
  const rows = proposal.rows.slice(0, PREVIEW_ROWS)
  const width = Math.min(PREVIEW_COLUMNS, Math.max(...rows.map((cells) => cells.length)))
  const columns = Array.from({ length: width }, (_, offset) => proposal.column + offset)
  const hidden = proposal.rows.length - rows.length
  return (
    <>
      <table className="sheet-proposal__preview" aria-label={t('ai.mcp.sheetProposal.preview')}>
        <thead>
          <tr>
            <th />
            {columns.map((column) => (
              <th key={column}>{columnName(column)}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((cells, offset) => (
            <tr key={offset}>
              <th>{proposal.row + offset + 1}</th>
              {columns.map((column) => (
                <td key={column}>{cells[column - proposal.column] ?? ''}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {hidden > 0 && (
        <span className="banner__detail">{t('ai.mcp.sheetProposal.moreRows', { count: hidden })}</span>
      )}
    </>
  )
}

/** The oldest write a client of the MCP server proposed; nothing changes until the person applies it. */
export function SheetMcpProposal(): React.JSX.Element | null {
  const t = useT()
  const generation = useWorkspace((state) => state.generation)
  const workbook = useWorkspace((state) => state.workbook)
  const readOnly = useWorkspace((state) => state.readOnly)
  const selectSheet = useWorkspace((state) => state.selectSheet)
  const updateSheet = useWorkspace((state) => state.updateSheet)
  const all = useSheetMcpProposals((state) => state.pending)
  const pending = useMemo(() => all.filter((item) => item.generation === generation), [all, generation])
  const proposal = pending[0]
  if (proposal === undefined || workbook === null) return null

  const index = workbook.sheets.findIndex((sheet) => sheet.name === proposal.sheet)
  const sheet = workbook.sheets[index]
  const target = targetOf(proposal)
  const filled = sheet === undefined ? 0 : filledIn(sheet, target)
  const apply = (): void => {
    if (sheet === undefined) return
    if (index !== workbook.activeSheet) selectSheet(index)
    updateSheet(pasteText(sheet, proposal.row, proposal.column, proposal.rows))
    discardSheetProposal(proposal.id)
  }

  return (
    <div
      className="banner banner--recovery sheet-proposal"
      role="region"
      aria-label={t('ai.mcp.sheetProposal.label')}
    >
      <div className="banner__text">
        <strong>
          {t('ai.mcp.sheetProposal.title', {
            author: proposal.author,
            range: `${quoteSheet(proposal.sheet)}!${describeRange(target)}`,
          })}
          {pending.length > 1 && ` (${t('ai.mcp.sheetProposal.position', { count: pending.length })})`}
        </strong>
        {proposal.reason !== undefined && <span className="banner__detail">{proposal.reason}</span>}
        <Preview proposal={proposal} />
        {sheet === undefined ? (
          <span className="banner__detail">
            {t('ai.mcp.sheetProposal.noSheet', { sheet: proposal.sheet })}
          </span>
        ) : (
          filled > 0 && (
            <span className="banner__detail">{t('ai.mcp.sheetProposal.replaces', { count: filled })}</span>
          )
        )}
      </div>
      <div className="banner__actions">
        <button
          type="button"
          className="banner__action"
          disabled={readOnly || sheet === undefined}
          onClick={apply}
        >
          {t('ai.mcp.sheetProposal.apply')}
        </button>
        <button
          type="button"
          className="banner__action banner__action--quiet"
          onClick={() => discardSheetProposal(proposal.id)}
        >
          {t('ai.mcp.sheetProposal.discard')}
        </button>
      </div>
    </div>
  )
}
