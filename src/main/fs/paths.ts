import type { Stats } from 'node:fs'
import { stat } from 'node:fs/promises'
import { isAbsolute, resolve } from 'node:path'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { isPdfPath, isSupportedExtension } from '@services/file/formats.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/** Keeps a huge file from freezing the UI. */
export const MAX_FILE_BYTES = 20 * 1024 * 1024

/**
 * The renderer **does not choose** the file main reads or writes: only a path picked in a native
 * dialog enters here. Lives as long as the app is open.
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

/**
 * PDFs main opened, which the viewer may show. Never a destination for content: only the
 * properties update `pdf-metadata.ts` builds is appended to them.
 */
const viewablePaths = new Set<string>()

export function allowViewing(path: string): string {
  const normalized = normalizePath(path)
  viewablePaths.add(normalized)
  return normalized
}

export function assertViewingAllowed(path: string): string {
  const normalized = normalizePath(path)
  if (!viewablePaths.has(normalized)) {
    throw new AppError(ErrorCode.PathNotAuthorized, t('errors.paths.unauthorized'))
  }
  return normalized
}

/** For the tests. */
export function resetAuthorizedPaths(): void {
  authorizedPaths.clear()
  viewablePaths.clear()
}

export async function assertReadableFile(path: string): Promise<void> {
  assertAbsolute(path)
  if (!isSupportedExtension(path)) {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.paths.unsupportedType'))
  }

  const info = await statFile(path)
  if (info.size > MAX_FILE_BYTES) {
    const limit = Math.round(MAX_FILE_BYTES / (1024 * 1024))
    throw new AppError(ErrorCode.FileTooLarge, t('errors.paths.fileTooLarge', { limit }))
  }
}

/** No size limit: Chromium streams the PDF, and its bytes never cross to the renderer. */
export async function assertViewablePdf(path: string): Promise<void> {
  assertAbsolute(path)
  if (!isPdfPath(path)) {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.paths.unsupportedType'))
  }
  await statFile(path)
}

function assertAbsolute(path: string): void {
  if (!isAbsolute(path)) {
    throw new AppError(ErrorCode.InvalidRequest, t('errors.paths.invalidPath'))
  }
}

async function statFile(path: string): Promise<Stats> {
  let info
  try {
    info = await stat(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  if (!info.isFile()) {
    throw new AppError(ErrorCode.NotAFile, t('errors.paths.notAFile'))
  }
  return info
}
