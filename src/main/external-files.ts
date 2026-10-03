import { extname, resolve } from 'node:path'
import type { BrowserWindow } from 'electron'
import { MenuCommand } from '@shared/types.js'
import { normalizePath } from './fs/paths.js'
import { sendMenuCommand } from './window.js'

const requested = new Set<string>()
const pending: string[] = []
let target: BrowserWindow | null = null

/**
 * Só os argumentos de abertura nativos dão acesso a arquivo fora dos recentes.
 * Um `.dotx` ou `.dotm` pela linha de comando cria um documento novo, como no Word.
 */
export function docxFromArguments(args: readonly string[], cwd: string): string | undefined {
  const path = args.find(
    (arg) => !arg.startsWith('-') && ['.docx', '.dotx', '.dotm'].includes(extname(arg).toLowerCase()),
  )
  return path === undefined ? undefined : resolve(cwd, path)
}

export function requestExternalFile(path: string): void {
  const normalized = normalizePath(path)
  requested.add(normalized)
  pending.push(normalized)
  flush()
}

export function isExternalFileRequested(path: string): boolean {
  return requested.has(normalizePath(path))
}

export function forgetExternalFileRequest(path: string): void {
  requested.delete(normalizePath(path))
}

export function externalFilesReady(window: BrowserWindow): void {
  target = window
  flush()
}

function flush(): void {
  if (target === null || target.isDestroyed()) return
  for (const path of pending.splice(0)) {
    sendMenuCommand(target, { command: MenuCommand.OpenRecent, path })
  }
}
