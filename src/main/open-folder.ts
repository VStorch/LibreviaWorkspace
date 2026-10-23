import { mkdir } from 'node:fs/promises'
import { shell } from 'electron'
import { logWarning } from './log.js'

export async function openFolder(folder: string): Promise<void> {
  await mkdir(folder, { recursive: true })
  // `openPath` returns the error message instead of throwing.
  const problem = await shell.openPath(folder)
  if (problem !== '') logWarning('folder', problem)
}
