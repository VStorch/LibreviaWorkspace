import { extname, resolve } from 'node:path'
import type { BrowserWindow } from 'electron'
import { MenuCommand } from '@shared/types.js'
import { normalizePath } from './fs/paths.js'
import { sendMenuCommand } from './window.js'

const requested = new Set<string>()
const pending: string[] = []
let target: BrowserWindow | null = null

/** Only native launch arguments can grant access to files outside recents. */
export function docxFromArguments(args: readonly string[], cwd: string): string | undefined {
  const path = args.find((arg) => !arg.startsWith('-') && extname(arg).toLowerCase() === '.docx')
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
