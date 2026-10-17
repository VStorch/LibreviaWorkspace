import { useId, useMemo } from 'react'
import {
  applyFill,
  previewFormula,
  readFillAnswer,
  readFormulaAnswer,
  type FillValue,
} from '@services/ai/sheet-answers.js'
import { cellsIn, writeText } from '@services/spreadsheet/edit.js'
import { formatCell } from '@services/spreadsheet/format.js'
import { cellRef, getCell, type Sheet, type WorkbookModel } from '@services/spreadsheet/model.js'
import { AiSheetTask } from '@shared/types.js'
import { useT } from '../i18n.js'
import { AiMarkdownPreview } from '../document/AiMarkdownPreview.js'
import type { SheetAiRun } from './useSheetAi.js'

interface ResultProps {
  readonly run: SheetAiRun
  readonly workbook: WorkbookModel
  readonly readOnly: boolean
  readonly onApply: (sheet: Sheet) => void
}

/** The explanation is read as it arrives; a formula or the values only once complete and checked. */
export function SheetAiResult(props: ResultProps): React.JSX.Element {
  const t = useT()
  const { run } = props
  if (run.task === AiSheetTask.Explain) {
    return (
      <div className="popover__field">
        <span>{t('ai.panel.result')}</span>
        <AiMarkdownPreview markdown={run.text} label={t('ai.panel.result')} />
      </div>
    )
  }
  if (run.phase !== 'done') {
    return (
      <label className="popover__field">
        <span>{t('ai.panel.result')}</span>
        <textarea className="ai-pane__output" readOnly rows={4} value={run.text} />
      </label>
    )
  }
  return run.task === AiSheetTask.Formula ? <FormulaPreview {...props} /> : <FillPreview {...props} />
}

/** Writing goes to the sheet the question was about: on another tab it would land in the wrong place. */
function useTargetSheet({ run, workbook, readOnly }: ResultProps): {
  readonly sheet: Sheet | undefined
  readonly blocked: boolean
  readonly elsewhere: string | null
} {
  const sheet = workbook.sheets[run.sheetIndex]
  const elsewhere = workbook.activeSheet === run.sheetIndex ? null : (sheet?.name ?? '')
  return { sheet, blocked: readOnly || elsewhere !== null || sheet === undefined, elsewhere }
}

function ElsewhereNote({ name }: { readonly name: string | null }): React.JSX.Element | null {
  const t = useT()
  if (name === null) return null
  return <p className="ai-pane__note">{t('ai.sheet.otherSheet', { sheet: name })}</p>
}

function FormulaPreview(props: ResultProps): React.JSX.Element {
  const t = useT()
  const { run, workbook, onApply } = props
  const { sheet, blocked, elsewhere } = useTargetSheet(props)
  const labelId = useId()
  const answer = useMemo(() => readFormulaAnswer(run.text), [run.text])
  const { fromRow: row, fromColumn: column } = run.target
  const ref = cellRef(row, column)
  const value = useMemo(
    () =>
      answer.kind === 'formula'
        ? formatCell(previewFormula(workbook, run.sheetIndex, { row, column }, answer.formula))
        : '',
    [answer, workbook, run.sheetIndex, row, column],
  )
  if (answer.kind === 'missing') {
    return (
      <p className="ai-pane__warning" role="alert">
        {t('ai.sheet.noFormulaInAnswer')}
      </p>
    )
  }
  return (
    <>
      <div className="popover__field">
        <span id={labelId}>{t('ai.sheet.formulaLabel')}</span>
        <code className="ai-sheet__formula" aria-labelledby={labelId}>
          {answer.formula}
        </code>
      </div>
      {answer.kind === 'invalid' ? (
        <p className="ai-pane__warning" role="alert">
          {t('ai.sheet.invalidFormula', { problem: answer.message })}
        </p>
      ) : (
        <p className="ai-pane__note" aria-label={t('ai.sheet.valueIn', { cell: ref })}>
          {t('ai.sheet.valueIn', { cell: ref })}: <strong>{value}</strong>
        </p>
      )}
      <ElsewhereNote name={elsewhere} />
      <div className="ai-pane__actions">
        <button
          type="button"
          className="btn btn--primary"
          disabled={answer.kind !== 'formula' || blocked}
          onClick={() => {
            if (sheet !== undefined && answer.kind === 'formula')
              onApply(writeText(sheet, row, column, answer.formula))
          }}
        >
          {t('ai.sheet.insert')}
        </button>
      </div>
    </>
  )
}

function shownValue(value: FillValue): string {
  return value === null ? '' : formatCell({ value })
}

function FillPreview(props: ResultProps): React.JSX.Element {
  const t = useT()
  const { run, onApply } = props
  const { sheet, blocked, elsewhere } = useTargetSheet(props)
  const rowCount = run.target.toRow - run.target.fromRow + 1
  const values = useMemo(() => readFillAnswer(run.text, rowCount), [run.text, rowCount])
  if (values === null || sheet === undefined) {
    return (
      <p className="ai-pane__warning" role="alert">
        {t('ai.sheet.badFormat')}
      </p>
    )
  }
  const cells = [...cellsIn(run.target)]
  return (
    <>
      <div className="ai-sheet__preview">
        <table aria-label={t('ai.sheet.preview')}>
          <thead>
            <tr>
              <th>{t('ai.sheet.cell')}</th>
              <th>{t('ai.sheet.current')}</th>
              <th>{t('ai.sheet.new')}</th>
            </tr>
          </thead>
          <tbody>
            {cells.map(({ row, column }, index) => (
              <tr key={row}>
                <td>{cellRef(row, column)}</td>
                <td>{formatCell(getCell(sheet, row, column))}</td>
                <td>{shownValue(values[index]!)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <ElsewhereNote name={elsewhere} />
      <div className="ai-pane__actions">
        <button
          type="button"
          className="btn btn--primary"
          disabled={blocked}
          onClick={() => onApply(applyFill(sheet, run.target, values))}
        >
          {t('ai.panel.apply')}
        </button>
      </div>
    </>
  )
}
