import { readFile, stat } from 'node:fs/promises'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { MAX_IMAGE_BYTES, detectImageMimeType, isImageWithinSizeLimit } from '@services/file/image.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/** Por assinatura de bytes, e não por extensão: a imagem viaja dentro do documento. */
export async function readImageAsDataUrl(path: string): Promise<string> {
  let size: number
  try {
    size = (await stat(path)).size
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  if (!isImageWithinSizeLimit(size)) {
    const limit = Math.round(MAX_IMAGE_BYTES / (1024 * 1024))
    throw new AppError(ErrorCode.FileTooLarge, t('errors.image.imageTooLarge', { limit }))
  }

  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  const mimeType = detectImageMimeType(bytes)
  if (mimeType === null) {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.image.unsupportedImage'))
  }

  return `data:${mimeType};base64,${bytes.toString('base64')}`
}
