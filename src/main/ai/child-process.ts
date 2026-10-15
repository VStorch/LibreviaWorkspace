import { spawn as nodeSpawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Readable } from 'node:stream'

const KILL_GRACE_MS = 2_000
const MAX_STDERR_LENGTH = 2_000

export type Spawn = (
  command: string,
  args: readonly string[],
  options: { cwd: string },
) => ChildProcessWithoutNullStreams

export interface WorkFolder {
  readonly makeTempDir?: () => Promise<string>
  readonly removeDir?: (dir: string) => Promise<void>
}

const IS_WINDOWS = process.platform === 'win32'

/**
 * Without a shell: arguments reach the program as they are, never interpreted. On POSIX the program
 * leads a process group of its own, so stopping it reaches the helpers it started.
 */
export const spawnWithoutShell: Spawn = (command, args, options) =>
  nodeSpawn(command, args, {
    cwd: options.cwd,
    shell: false,
    detached: !IS_WINDOWS,
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  })

/** An empty folder of its own, so the program has no project of the person's to read. */
export function makeWorkFolder(folder: WorkFolder): Promise<string> {
  return (folder.makeTempDir ?? (() => mkdtemp(join(tmpdir(), 'librevia-ai-'))))()
}

export function removeWorkFolder(folder: WorkFolder, dir: string): Promise<void> {
  return (folder.removeDir ?? ((path) => rm(path, { recursive: true, force: true })))(dir)
}

export function exitOf(
  child: ChildProcessWithoutNullStreams,
  failedToStart: (cause: Error) => Error,
): Promise<number | null> {
  return new Promise((resolve, reject) => {
    child.once('error', (cause) => reject(failedToStart(cause)))
    child.once('close', (code) => resolve(code))
  })
}

export function collect(stream: Readable): () => string {
  let text = ''
  stream.on('data', (chunk: Buffer | string) => {
    text = (text + chunk.toString()).slice(-MAX_STDERR_LENGTH)
  })
  return () => text
}

export interface ProcessTreeKiller {
  readonly windows: boolean
  readonly killGroup: (pid: number, signal: NodeJS.Signals) => void
  readonly runTaskkill: (args: readonly string[]) => void
}

const systemKiller: ProcessTreeKiller = {
  windows: IS_WINDOWS,
  killGroup: (pid, signal) => process.kill(-pid, signal),
  runTaskkill: (args) => {
    nodeSpawn('taskkill', args, { shell: false, stdio: 'ignore', windowsHide: true }).on(
      'error',
      () => undefined,
    )
  },
}

/** The program and every process it started; the program alone if its group is gone. */
export function killTree(
  child: ChildProcessWithoutNullStreams,
  signal: NodeJS.Signals,
  killer: ProcessTreeKiller = systemKiller,
): void {
  const pid = child.pid
  if (pid === undefined) {
    child.kill(signal)
    return
  }
  if (killer.windows) {
    killer.runTaskkill(['/pid', String(pid), '/T', '/F'])
    return
  }
  try {
    killer.killGroup(pid, signal)
  } catch {
    child.kill(signal)
  }
}

function hasEnded(child: ChildProcessWithoutNullStreams): boolean {
  return child.exitCode !== null || child.signalCode !== null
}

export function terminate(child: ChildProcessWithoutNullStreams, killer?: ProcessTreeKiller): void {
  if (hasEnded(child)) return
  killTree(child, 'SIGTERM', killer)
  const timer = setTimeout(() => killTree(child, 'SIGKILL', killer), KILL_GRACE_MS)
  timer.unref()
  child.once('exit', () => clearTimeout(timer))
}

/** On quit there is no later: a timer would never fire. */
export function terminateNow(child: ChildProcessWithoutNullStreams, killer?: ProcessTreeKiller): void {
  if (!hasEnded(child)) killTree(child, 'SIGKILL', killer)
}
