/**
 * Erros com mensagem destinada ao usuário final.
 *
 * Regra: nada de stack trace, caminho absoluto ou detalhe interno atravessa o
 * IPC. O renderer recebe um código estável e uma frase compreensível; o
 * diagnóstico técnico fica no log do processo main.
 */

export const ErrorCode = {
  InvalidRequest: 'INVALID_REQUEST',
  UnknownChannel: 'UNKNOWN_CHANNEL',
  Internal: 'INTERNAL',
} as const

export type ErrorCode = (typeof ErrorCode)[keyof typeof ErrorCode]

export interface SerializedError {
  readonly code: ErrorCode
  /** Frase pronta para exibição, em português, sem jargão. */
  readonly message: string
  /** Detalhe opcional já higienizado (ex.: qual campo falhou na validação). */
  readonly detail?: string
}

export class AppError extends Error {
  readonly code: ErrorCode
  readonly detail: string | undefined

  constructor(code: ErrorCode, message: string, detail?: string) {
    super(message)
    this.name = 'AppError'
    this.code = code
    this.detail = detail
  }

  toSerialized(): SerializedError {
    return this.detail === undefined
      ? { code: this.code, message: this.message }
      : { code: this.code, message: this.message, detail: this.detail }
  }
}

/** Converte qualquer valor lançado num erro seguro para cruzar o IPC. */
export function toSerializedError(cause: unknown): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return {
    code: ErrorCode.Internal,
    message: 'Ocorreu um erro inesperado. A operação não foi concluída.',
  }
}
