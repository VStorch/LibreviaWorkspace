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

/** Without a shell: arguments reach the program as they are, never interpreted. */
export const spawnWithoutShell: Spawn = (command, args, options) =>
  nodeSpawn(command, args, {
    cwd: options.cwd,
    shell: false,
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

export function terminate(child: ChildProcessWithoutNullStreams): void {
  if (child.exitCode !== null || child.signalCode !== null) return
  child.kill('SIGTERM')
  const timer = setTimeout(() => child.kill('SIGKILL'), KILL_GRACE_MS)
  timer.unref()
  child.once('exit', () => clearTimeout(timer))
}
