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

  FileNotFound: 'FILE_NOT_FOUND',
  NotAFile: 'NOT_A_FILE',
  FileTooLarge: 'FILE_TOO_LARGE',
  PermissionDenied: 'PERMISSION_DENIED',
  UnsupportedFormat: 'UNSUPPORTED_FORMAT',
  PathNotAuthorized: 'PATH_NOT_AUTHORIZED',
  ReadFailed: 'READ_FAILED',
  WriteFailed: 'WRITE_FAILED',
  NotTextFile: 'NOT_TEXT_FILE',

  /** O serviço de formatos não pôde ser iniciado — instalação incompleta. */
  SidecarUnavailable: 'SIDECAR_UNAVAILABLE',
  /** Demorou além do limite. O documento aberto continua intacto. */
  SidecarTimeout: 'SIDECAR_TIMEOUT',
  /** Morreu no meio da operação, ou respondeu algo que não entendemos. */
  SidecarFailed: 'SIDECAR_FAILED',
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

/**
 * Traduz o `errno` do sistema de arquivos numa frase que o usuário entenda.
 *
 * Sem isto, uma pasta de rede fora do ar produz "EBUSY" na tela — que não diz
 * nada a quem só quer saber se pode continuar trabalhando.
 */
export function fromFileSystemError(cause: unknown, operation: 'leitura' | 'escrita'): AppError {
  const code = typeof cause === 'object' && cause !== null ? (cause as { code?: string }).code : undefined

  switch (code) {
    case 'ENOENT':
      return new AppError(
        ErrorCode.FileNotFound,
        'O arquivo não foi encontrado. Ele pode ter sido movido ou excluído.',
      )
    case 'EACCES':
    case 'EPERM':
      return new AppError(
        ErrorCode.PermissionDenied,
        'Você não tem permissão para acessar este arquivo. Verifique com quem administra a pasta.',
      )
    case 'EISDIR':
      return new AppError(ErrorCode.NotAFile, 'O caminho indicado é uma pasta, não um arquivo.')
    case 'EROFS':
      return new AppError(
        ErrorCode.WriteFailed,
        'Este local é somente leitura. Salve o arquivo em outra pasta.',
      )
    case 'ENOSPC':
      return new AppError(ErrorCode.WriteFailed, 'Não há espaço em disco para salvar o arquivo.')
    case 'EBUSY':
      return new AppError(
        ErrorCode.WriteFailed,
        'O arquivo está em uso por outro programa. Feche-o e tente novamente.',
      )
    // Típicos de pasta de rede que caiu no meio da operação.
    case 'ENETDOWN':
    case 'ENETUNREACH':
    case 'EHOSTDOWN':
    case 'EHOSTUNREACH':
    case 'ESTALE':
    case 'ETIMEDOUT':
      return new AppError(
        operation === 'leitura' ? ErrorCode.ReadFailed : ErrorCode.WriteFailed,
        'A pasta de rede não respondeu. Verifique a conexão e tente novamente.',
      )
    default:
      return new AppError(
        operation === 'leitura' ? ErrorCode.ReadFailed : ErrorCode.WriteFailed,
        operation === 'leitura'
          ? 'Não foi possível ler o arquivo.'
          : 'Não foi possível salvar o arquivo. O conteúdo original foi preservado.',
      )
  }
}
