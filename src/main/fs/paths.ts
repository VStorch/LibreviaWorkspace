import { stat } from 'node:fs/promises'
import { isAbsolute, resolve } from 'node:path'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { isSupportedExtension } from '@services/file/formats.js'

/** Teto de leitura da Fase 1. Protege contra travar a interface com um arquivo enorme. */
export const MAX_FILE_BYTES = 20 * 1024 * 1024

/**
 * Autorização de caminhos.
 *
 * O renderer é tratado como não confiável, então ele **não escolhe** que
 * arquivo o processo main lê ou grava. Um caminho só entra neste conjunto
 * quando o próprio usuário o escolheu num diálogo nativo — e é por isso que
 * `file:save` recebe um caminho do renderer sem que isso seja uma brecha: se
 * o caminho não estiver aqui, a operação é recusada.
 *
 * O conjunto vive só enquanto o aplicativo estiver aberto.
 */
const authorizedPaths = new Set<string>()

export function normalizePath(path: string): string {
  return resolve(path)
}

export function authorizePath(path: string): string {
  const normalized = normalizePath(path)
  authorizedPaths.add(normalized)
  return normalized
}

export function isPathAuthorized(path: string): boolean {
  return authorizedPaths.has(normalizePath(path))
}

export function assertPathAuthorized(path: string): string {
  const normalized = normalizePath(path)
  if (!authorizedPaths.has(normalized)) {
    throw new AppError(
      ErrorCode.PathNotAuthorized,
      'Esta operação foi recusada porque o arquivo não foi aberto nem escolhido por você nesta sessão.',
    )
  }
  return normalized
}

/** Apenas para testes: devolve o conjunto ao estado inicial. */
export function resetAuthorizedPaths(): void {
  authorizedPaths.clear()
}

/** Valida que o caminho é legível, é um arquivo comum e cabe no limite. */
export async function assertReadableFile(path: string): Promise<void> {
  if (!isAbsolute(path)) {
    throw new AppError(ErrorCode.InvalidRequest, 'O caminho do arquivo é inválido.')
  }

  if (!isSupportedExtension(path)) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este tipo de arquivo ainda não é suportado. Por enquanto o aplicativo abre arquivos .txt.',
    )
  }

  let info
  try {
    info = await stat(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  if (!info.isFile()) {
    throw new AppError(ErrorCode.NotAFile, 'O caminho indicado não é um arquivo.')
  }

  if (info.size > MAX_FILE_BYTES) {
    const limit = Math.round(MAX_FILE_BYTES / (1024 * 1024))
    throw new AppError(
      ErrorCode.FileTooLarge,
      `O arquivo é grande demais para ser aberto (limite atual: ${limit} MB).`,
    )
  }
}
