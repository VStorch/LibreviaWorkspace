import { readFile } from 'node:fs/promises'
import { fromFileSystemError } from '@shared/errors.js'
import { editorPreferences } from '../preferences.js'

/** A file opened by the user: a file system failure becomes a sentence in the interface language. */
export async function readFileBytes(path: string): Promise<Buffer> {
  try {
    return await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'read', editorPreferences().language)
  }
}
