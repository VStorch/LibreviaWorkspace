import { useRef, useState } from 'react'
import { getCell, type Sheet } from '@services/spreadsheet/model.js'
import { describeRange, type Range } from '@services/spreadsheet/edit.js'
import { checkFormula } from '@services/spreadsheet/formula/validate.js'
import { formatCell } from '@services/spreadsheet/format.js'
import { useT } from '../i18n.js'

/**
 * A célula mostra o resultado, e aqui se vê a fórmula. A fórmula é conferida
 * antes de entrar: `=ABS(1;2)` recebe a frase, e não o `#VALOR!`.
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
  // O valor cru, e não o formatado: reeditar "R$ 1.234,50" devolveria texto.
  const stored = cell?.formula ?? (cell?.value === undefined ? '' : String(cell.value))

  const [draft, setDraft] = useState(stored)
  const [problem, setProblem] = useState<string | null>(null)
  const input = useRef<HTMLInputElement>(null)

  /** Só a troca de célula recarrega a barra: o recálculo apagaria o que se está escrevendo. */
  const anchor = describeRange(range)
  const loaded = useRef(anchor)
  if (loaded.current !== anchor) {
    loaded.current = anchor
    // Na renderização, e não num efeito, para não piscar um valor velho.
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
        // Sair confirma, como no Excel, menos com problema, que descartaria o que se escreveu.
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
