import { stat } from 'node:fs/promises'
import { isAbsolute, resolve } from 'node:path'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { isSupportedExtension } from '@services/file/formats.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/** Teto de leitura: protege contra travar a interface com um arquivo enorme. */
export const MAX_FILE_BYTES = 20 * 1024 * 1024

/**
 * O renderer **não escolhe** o arquivo que o main lê ou grava: só entra aqui o
 * caminho escolhido num diálogo nativo. Vive enquanto o aplicativo estiver aberto.
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
    throw new AppError(ErrorCode.PathNotAuthorized, t('errors.paths.unauthorized'))
  }
  return normalized
}

/** Para os testes. */
export function resetAuthorizedPaths(): void {
  authorizedPaths.clear()
}

/** Valida que o caminho é legível, é um arquivo comum e cabe no limite. */
export async function assertReadableFile(path: string): Promise<void> {
  if (!isAbsolute(path)) {
    throw new AppError(ErrorCode.InvalidRequest, t('errors.paths.invalidPath'))
  }

  if (!isSupportedExtension(path)) {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.paths.unsupportedType'))
  }

  let info
  try {
    info = await stat(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  if (!info.isFile()) {
    throw new AppError(ErrorCode.NotAFile, t('errors.paths.notAFile'))
  }

  if (info.size > MAX_FILE_BYTES) {
    const limit = Math.round(MAX_FILE_BYTES / (1024 * 1024))
    throw new AppError(ErrorCode.FileTooLarge, t('errors.paths.fileTooLarge', { limit }))
  }
}
