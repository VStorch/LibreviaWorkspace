/**
 * **Valores**, e não exceções: `=A1/0` produz `#DIV/0!` e não derruba o
 * recálculo. Os nomes são os do Excel em português.
 */

export const FormulaError = {
  Div0: '#DIV/0!',
  Value: '#VALOR!',
  Ref: '#REF!',
  Name: '#NOME?',
  Num: '#NÚM!',
  NA: '#N/D',
  /** O Excel não tem este erro: mostra zero e um aviso, e um zero silencioso é resultado errado. */
  Circular: '#CIRC!',
} as const
export type FormulaError = (typeof FormulaError)[keyof typeof FormulaError]

const ALL: readonly string[] = Object.values(FormulaError)

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'string' && ALL.includes(value)
}

export class ParseError extends Error {
  constructor(
    message: string,
    readonly position: number,
  ) {
    super(message)
    this.name = 'ParseError'
  }
}
