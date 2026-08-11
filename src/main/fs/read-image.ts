import { readFile, stat } from 'node:fs/promises'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { MAX_IMAGE_BYTES, detectImageMimeType, isImageWithinSizeLimit } from '@services/file/image.js'

/**
 * Lê uma imagem do disco e devolve um data URI.
 *
 * A validação é por assinatura de bytes, não por extensão: o arquivo vai ser
 * embutido no documento do usuário e distribuído junto com ele, então o que
 * entra precisa ser mesmo uma imagem de formato conhecido.
 */
export async function readImageAsDataUrl(path: string): Promise<string> {
  let size: number
  try {
    size = (await stat(path)).size
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  if (!isImageWithinSizeLimit(size)) {
    const limit = Math.round(MAX_IMAGE_BYTES / (1024 * 1024))
    throw new AppError(
      ErrorCode.FileTooLarge,
      `A imagem é grande demais para ser inserida (limite: ${limit} MB).`,
    )
  }

  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  const mimeType = detectImageMimeType(bytes)
  if (mimeType === null) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este arquivo não é uma imagem suportada. Use PNG, JPEG, GIF ou WebP.',
    )
  }

  return `data:${mimeType};base64,${bytes.toString('base64')}`
}
