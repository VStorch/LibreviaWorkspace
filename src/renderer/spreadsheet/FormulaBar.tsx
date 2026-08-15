import { useRef, useState } from 'react'
import { getCell, type Sheet } from '@services/spreadsheet/model.js'
import { describeRange, type Range } from '@services/spreadsheet/edit.js'
import { checkFormula } from '@services/spreadsheet/formula/validate.js'
import { formatCell } from '@services/spreadsheet/format.js'

/**
 * Barra de fórmulas.
 *
 * Existe porque a célula mostra o **resultado**, e o usuário precisa de um
 * lugar que mostre a **fórmula**. Sem ela, a única forma de reler uma conta
 * seria entrar no modo de edição da célula — e sair dele sem querer alteraria o
 * conteúdo.
 *
 * A fórmula digitada é conferida antes de entrar: quem escreveu `=ABS(1;2)`
 * recebe a frase em vez de descobrir pelo `#VALOR!` na célula que o problema
 * era o número de argumentos.
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
  const cell = getCell(sheet, range.fromRow, range.fromColumn)
  // O que a célula guarda: a fórmula quando há, e o valor cru quando não.
  // Valor cru, e não formatado: reeditar "R$ 1.234,50" devolveria texto.
  const stored = cell?.formula ?? (cell?.value === undefined ? '' : String(cell.value))

  const [draft, setDraft] = useState(stored)
  const [problem, setProblem] = useState<string | null>(null)
  const input = useRef<HTMLInputElement>(null)

  /**
   * Trocar de célula recarrega a barra; recalcular, não.
   *
   * O efeito depende só da referência selecionada, e de propósito: se ele
   * seguisse o conteúdo, cada recálculo disparado pela digitação em outra
   * célula apagaria o que o usuário está escrevendo aqui.
   */
  const anchor = describeRange(range)
  const loaded = useRef(anchor)
  if (loaded.current !== anchor) {
    loaded.current = anchor
    // Durante a renderização, e não num efeito: o campo já aparece com o
    // conteúdo certo, sem o piscar de um valor velho por um quadro.
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
      <span className="formula-bar__ref" title="Célula selecionada">
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
        aria-label="Fórmula ou conteúdo da célula"
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
        // Sair do campo confirma, como no Excel — menos quando há problema, que
        // descartaria o que o usuário escreveu sem ele ter chance de corrigir.
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
