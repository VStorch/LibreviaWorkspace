import { appendFileSync, mkdirSync, renameSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { inspect } from 'node:util'

export const MAX_LOG_BYTES = 1024 * 1024
export const LOG_FILE = 'librevia.log'
export const PREVIOUS_LOG_FILE = 'librevia.old.log'

const INSPECT_DEPTH = 4

export const LogLevel = {
  Info: 'INFO',
  Warning: 'WARN',
  Error: 'ERROR',
} as const

export type LogLevel = (typeof LogLevel)[keyof typeof LogLevel]

let folder: string | null = null

export function useLogFolder(path: string): void {
  folder = path
}

export function logFolder(): string | null {
  return folder
}

export function logFilePath(): string | null {
  return folder === null ? null : join(folder, LOG_FILE)
}

export function formatLogLine(level: LogLevel, scope: string, parts: readonly unknown[], now: Date): string {
  return `${now.toISOString()} ${level} [${scope}] ${parts.map(describe).join(' ')}\n`
}

function describe(part: unknown): string {
  if (typeof part === 'string') return part
  if (part instanceof Error) return part.stack ?? `${part.name}: ${part.message}`
  return inspect(part, { depth: INSPECT_DEPTH, breakLength: Infinity })
}

function rotateIfFull(directory: string, file: string): void {
  let size: number
  try {
    size = statSync(file).size
  } catch {
    return
  }
  if (size >= MAX_LOG_BYTES) renameSync(file, join(directory, PREVIOUS_LOG_FILE))
}

function append(level: LogLevel, scope: string, parts: readonly unknown[]): void {
  if (folder === null) return
  const file = join(folder, LOG_FILE)
  try {
    mkdirSync(folder, { recursive: true })
    rotateIfFull(folder, file)
    appendFileSync(file, formatLogLine(level, scope, parts, new Date()))
  } catch (cause) {
    console.error('[log] could not write the log file:', cause)
  }
}

export function logInfo(scope: string, ...parts: unknown[]): void {
  console.info(`[${scope}]`, ...parts)
  append(LogLevel.Info, scope, parts)
}

export function logWarning(scope: string, ...parts: unknown[]): void {
  console.warn(`[${scope}]`, ...parts)
  append(LogLevel.Warning, scope, parts)
}

export function logError(scope: string, ...parts: unknown[]): void {
  console.error(`[${scope}]`, ...parts)
  append(LogLevel.Error, scope, parts)
}
