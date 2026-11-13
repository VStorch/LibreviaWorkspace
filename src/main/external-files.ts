import { resolve } from 'node:path'
import type { BrowserWindow } from 'electron'
import { isOpenableExtension } from '@services/file/formats.js'
import type { MenuCommandPayload } from '@shared/api.js'
import { MenuCommand } from '@shared/types.js'
import { normalizePath } from './fs/paths.js'
import { sendMenuCommand } from './window.js'

const requested = new Set<string>()
const pending: string[] = []
let target: BrowserWindow | null = null
/** What a window opened for one file or command does as soon as its page can hear it. */
const firstCommands = new WeakMap<BrowserWindow, MenuCommandPayload>()

/**
 * The first openable path among the arguments, resolved against `cwd`. Only native open arguments
 * grant access to a file outside the recent list.
 */
export function fileFromArguments(args: readonly string[], cwd: string): string | undefined {
  const path = args.find((arg) => !arg.startsWith('-') && isOpenableExtension(arg))
  return path === undefined ? undefined : resolve(cwd, path)
}

export function requestExternalFile(path: string): void {
  pending.push(allowExternalFile(path))
  flush()
}

/** A path picked in the open dialog: the renderer may open it once, as a file manager request. */
export function allowExternalFile(path: string): string {
  const normalized = normalizePath(path)
  requested.add(normalized)
  return normalized
}

export function isExternalFileRequested(path: string): boolean {
  return requested.has(normalizePath(path))
}

export function forgetExternalFileRequest(path: string): void {
  requested.delete(normalizePath(path))
}

export function startWindowWith(window: BrowserWindow, command: MenuCommandPayload): void {
  firstCommands.set(window, command)
}

/** The window the file manager's files go to: the one that last became ready. */
export function externalFilesTarget(): BrowserWindow | null {
  return target === null || target.isDestroyed() ? null : target
}

export function externalFilesReady(window: BrowserWindow): void {
  const first = firstCommands.get(window)
  firstCommands.delete(window)
  if (first !== undefined) sendMenuCommand(window, first)
  target = window
  flush()
}

function flush(): void {
  if (target === null || target.isDestroyed()) return
  for (const path of pending.splice(0)) {
    sendMenuCommand(target, { command: MenuCommand.OpenRecent, path })
  }
}
