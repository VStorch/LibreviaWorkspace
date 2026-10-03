import { readFile } from 'node:fs/promises'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/** Com o BOM, para os arquivos do Bloco de Notas do Windows. Sem BOM, UTF-8. */
export async function readTextFile(path: string): Promise<string> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  if (startsWith(bytes, [0xef, 0xbb, 0xbf])) {
    return bytes.subarray(3).toString('utf8')
  }

  if (startsWith(bytes, [0xff, 0xfe])) {
    return bytes.subarray(2).toString('utf16le')
  }

  if (startsWith(bytes, [0xfe, 0xff])) {
    return swapByteOrder(bytes.subarray(2)).toString('utf16le')
  }

  assertNotBinary(bytes)
  return bytes.toString('utf8')
}

function startsWith(bytes: Buffer, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false
  return prefix.every((byte, index) => bytes[index] === byte)
}

/** Byte zero denuncia binário: abri-lo encheria o editor de lixo, e salvá-lo destruiria o arquivo. */
function assertNotBinary(bytes: Buffer): void {
  const sample = bytes.subarray(0, 8192)
  if (sample.includes(0)) {
    throw new AppError(ErrorCode.NotTextFile, t('errors.text.notTextFile'))
  }
}

function swapByteOrder(bytes: Buffer): Buffer {
  const swapped = Buffer.from(bytes)
  // `swap16` exige comprimento par.
  return swapped.subarray(0, swapped.length - (swapped.length % 2)).swap16()
}
