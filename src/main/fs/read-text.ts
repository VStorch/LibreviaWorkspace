import { readFile } from 'node:fs/promises'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/** With a BOM, for Windows Notepad files. Without one, UTF-8. */
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

/**
 * A zero byte gives a binary away: opening it would fill the editor with garbage, and saving it
 * would destroy the file.
 */
function assertNotBinary(bytes: Buffer): void {
  const sample = bytes.subarray(0, 8192)
  if (sample.includes(0)) {
    throw new AppError(ErrorCode.NotTextFile, t('errors.text.notTextFile'))
  }
}

function swapByteOrder(bytes: Buffer): Buffer {
  const swapped = Buffer.from(bytes)
  // `swap16` needs an even length.
  return swapped.subarray(0, swapped.length - (swapped.length % 2)).swap16()
}
