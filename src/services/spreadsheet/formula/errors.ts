/**
 * Valores de erro de fórmula.
 *
 * São **valores**, não exceções: no Excel `=A1/0` não interrompe nada, ela
 * produz `#DIV/0!`, e quem soma essa célula propaga o erro. Lançar exceção aqui
 * obrigaria cada função a capturar, e um erro em uma célula derrubaria o
 * recálculo da planilha inteira.
 *
 * Os nomes são os do Excel em português, porque é com eles que o usuário vai
 * comparar quando abrir a mesma planilha nos dois programas.
 */

export const FormulaError = {
  /** Divisão por zero. */
  Div0: '#DIV/0!',
  /** Tipo errado: texto onde se esperava número. */
  Value: '#VALOR!',
  /** Referência que não existe mais — a linha ou coluna foi excluída. */
  Ref: '#REF!',
  /** Nome de função desconhecido. */
  Name: '#NOME?',
  /** Número fora do domínio: raiz de negativo, argumento impossível. */
  Num: '#NÚM!',
  /** Valor não disponível: a procura não encontrou nada. */
  NA: '#N/D',
  /**
   * Referência circular.
   *
   * O Excel não tem esse erro: ele exibe zero e abre um aviso, e a planilha
   * segue com um número inventado no meio. Aqui a célula diz o que houve, porque
   * um zero silencioso numa fórmula circular é um resultado errado que ninguém
   * percebe.
   */
  Circular: '#CIRC!',
} as const
export type FormulaError = (typeof FormulaError)[keyof typeof FormulaError]

const ALL: readonly string[] = Object.values(FormulaError)

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'string' && ALL.includes(value)
}

/** Erro de escrita da fórmula, detectado antes de avaliar. */
export class ParseError extends Error {
  constructor(
    message: string,
    /** Posição na fórmula, para a mensagem apontar onde está o problema. */
    readonly position: number,
  ) {
    super(message)
    this.name = 'ParseError'
  }
}
