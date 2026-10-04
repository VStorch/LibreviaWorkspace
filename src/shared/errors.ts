/** Nada de stack trace, caminho absoluto ou detalhe interno atravessa o IPC: o detalhe fica no log do main. */

import { Language, translate, type MessageKey } from './i18n/index.js'

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
export function toSerializedError(cause: unknown, language: Language = Language.Portuguese): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return {
    code: ErrorCode.Internal,
    message: translate(language, 'errors.unexpected'),
  }
}

const FILE_SYSTEM_ERRORS: ReadonlyMap<string, readonly [ErrorCode, MessageKey]> = new Map([
  ['ENOENT', [ErrorCode.FileNotFound, 'errors.fs.fileNotFound']],
  ['EACCES', [ErrorCode.PermissionDenied, 'errors.fs.permissionDenied']],
  ['EPERM', [ErrorCode.PermissionDenied, 'errors.fs.permissionDenied']],
  ['EISDIR', [ErrorCode.NotAFile, 'errors.fs.notAFile']],
  ['EROFS', [ErrorCode.WriteFailed, 'errors.fs.readOnlyLocation']],
  ['ENOSPC', [ErrorCode.WriteFailed, 'errors.fs.diskFull']],
  // Diferente de disco cheio, e a diferença muda o que a pessoa faz: aqui o
  // disco tem espaço, mas a cota dela na pasta de rede acabou.
  ['EDQUOT', [ErrorCode.WriteFailed, 'errors.fs.quotaExceeded']],
  ['ENAMETOOLONG', [ErrorCode.WriteFailed, 'errors.fs.nameTooLong']],
  ['EBUSY', [ErrorCode.WriteFailed, 'errors.fs.fileInUse']],
])

/** Típicos de pasta de rede que caiu no meio da operação. */
const NETWORK_ERRORS: ReadonlySet<string> = new Set([
  'ENETDOWN',
  'ENETUNREACH',
  'EHOSTDOWN',
  'EHOSTUNREACH',
  'ESTALE',
  'ETIMEDOUT',
])

/** Sem isto, uma pasta de rede fora do ar mostraria "EBUSY" na tela. */
export function fromFileSystemError(
  cause: unknown,
  operation: 'leitura' | 'escrita',
  language: Language = Language.Portuguese,
): AppError {
  const code = typeof cause === 'object' && cause !== null ? (cause as { code?: string }).code : undefined
  const known = code === undefined ? undefined : FILE_SYSTEM_ERRORS.get(code)
  if (known !== undefined) return new AppError(known[0], translate(language, known[1]))

  const reading = operation === 'leitura'
  const failed = reading ? ErrorCode.ReadFailed : ErrorCode.WriteFailed
  if (code !== undefined && NETWORK_ERRORS.has(code)) {
    return new AppError(failed, translate(language, 'errors.fs.networkTimeout'))
  }
  return new AppError(failed, translate(language, reading ? 'errors.fs.readFailed' : 'errors.fs.writeFailed'))
}
