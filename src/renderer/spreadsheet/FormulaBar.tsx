import { useRef, useState } from 'react'
import { getCell, type Sheet } from '@services/spreadsheet/model.js'
import { describeRange, type Range } from '@services/spreadsheet/edit.js'
import { checkFormula } from '@services/spreadsheet/formula/validate.js'
import { formatCell } from '@services/spreadsheet/format.js'
import { useT } from '../i18n.js'

/**
 * The cell shows the result, and the formula shows here. The formula is checked before it goes in:
 * `=ABS(1;2)` gets the sentence, not `#VALOR!`.
 */
export function FormulaBar({
  sheet,
  range,
  onCommit,
}: {
  sheet: Sheet
  range: Range
  onCommit: (text: string) => void
}): React.JSX.Element {
  const t = useT()
  const cell = getCell(sheet, range.fromRow, range.fromColumn)
  // The raw value, not the formatted one: re-editing "R$ 1.234,50" would give back text.
  const stored = cell?.formula ?? (cell?.value === undefined ? '' : String(cell.value))

  const [draft, setDraft] = useState(stored)
  const [problem, setProblem] = useState<string | null>(null)
  const input = useRef<HTMLInputElement>(null)

  /** Only a cell change reloads the bar: recalculation would erase what is being typed. */
  const anchor = describeRange(range)
  const loaded = useRef(anchor)
  if (loaded.current !== anchor) {
    loaded.current = anchor
    // During render, not in an effect, so an old value does not flash.
    setDraft(stored)
    setProblem(null)
  }

  const commit = (): void => {
    if (draft.startsWith('=')) {
      const found = checkFormula(draft)
      if (found !== null) {
        setProblem(found.message)
        input.current?.focus()
        return
      }
    }

    setProblem(null)
    onCommit(draft)
  }

  return (
    <div className="formula-bar">
      <span className="formula-bar__ref" title={t('spreadsheet.formulaBar.selectedCell')}>
        {anchor}
      </span>

      <span className="formula-bar__fx" aria-hidden="true">
        ƒx
      </span>

      <input
        ref={input}
        className={problem === null ? 'formula-bar__input' : 'formula-bar__input formula-bar__input--bad'}
        value={draft}
        spellCheck={false}
        aria-label={t('spreadsheet.formulaBar.inputLabel')}
        aria-invalid={problem !== null}
        onChange={(event) => {
          setDraft(event.target.value)
          if (problem !== null) setProblem(null)
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault()
            commit()
          }
          if (event.key === 'Escape') {
            event.preventDefault()
            setDraft(stored)
            setProblem(null)
          }
        }}
        // Leaving confirms, as in Excel, except when there is a problem, which would discard what
        // was typed.
        onBlur={() => {
          if (problem === null && draft !== stored) commit()
        }}
      />

      {problem !== null && (
        <span className="formula-bar__problem" role="alert">
          {problem}
        </span>
      )}

      <span className="formula-bar__value" title="Resultado">
        {cell?.formula === undefined ? '' : formatCell(cell)}
      </span>
    </div>
  )
}
