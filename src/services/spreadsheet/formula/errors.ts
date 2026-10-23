import { translate, Language, type MessageKey, type Vars } from '@shared/i18n/index.js'

/**
 * **Values**, not exceptions: `=A1/0` produces `#DIV/0!` and does not bring recalculation down. The
 * names are Excel's in Portuguese.
 */

export const FormulaError = {
  Div0: '#DIV/0!',
  Value: '#VALOR!',
  Ref: '#REF!',
  Name: '#NOME?',
  Num: '#NÚM!',
  NA: '#N/D',
  /** Excel lacks this error: it shows zero and a warning, and a silent zero is a wrong result. */
  Circular: '#CIRC!',
} as const
export type FormulaError = (typeof FormulaError)[keyof typeof FormulaError]

const ALL: readonly string[] = Object.values(FormulaError)

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'string' && ALL.includes(value)
}

/**
 * Carries the catalog key, not the sentence: the formula bar and the MCP client read it in different
 * languages.
 */
export class ParseError extends Error {
  constructor(
    readonly key: MessageKey,
    readonly position: number,
    readonly vars?: Vars,
  ) {
    super(translate(Language.English, key, vars))
    this.name = 'ParseError'
  }

  describe(language: Language): string {
    return translate(language, this.key, this.vars)
  }
}
