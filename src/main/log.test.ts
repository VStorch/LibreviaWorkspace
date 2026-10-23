import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  LOG_FILE,
  LogLevel,
  MAX_LOG_BYTES,
  PREVIOUS_LOG_FILE,
  formatLogLine,
  logError,
  logWarning,
  useLogFolder,
} from './log.js'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-log-'))
  useLogFolder(directory)
  vi.spyOn(console, 'error').mockImplementation(() => undefined)
  vi.spyOn(console, 'warn').mockImplementation(() => undefined)
})

afterEach(async () => {
  vi.restoreAllMocks()
  await rm(directory, { recursive: true, force: true })
})

const readLog = (name = LOG_FILE) => readFile(join(directory, name), 'utf8')

describe('log file', () => {
  it('formats a line with time, level and scope', () => {
    const line = formatLogLine(LogLevel.Error, 'sidecar', ['exited', 3], new Date(Date.UTC(2026, 9, 23, 12)))

    expect(line).toBe('2026-10-23T12:00:00.000Z ERROR [sidecar] exited 3\n')
  })

  it('keeps the stack of an error, which is what tells where it failed', () => {
    const line = formatLogLine(LogLevel.Error, 'main', [new Error('boom')], new Date())

    expect(line).toContain('Error: boom\n    at ')
  })

  it('appends each message to the same file', async () => {
    logError('ipc', 'first')
    logWarning('ipc', 'second')

    const lines = (await readLog()).trimEnd().split('\n')
    expect(lines).toHaveLength(2)
    expect(lines[0]).toContain('ERROR [ipc] first')
    expect(lines[1]).toContain('WARN [ipc] second')
  })

  it('moves a full log aside and starts a new one', async () => {
    await writeFile(join(directory, LOG_FILE), 'x'.repeat(MAX_LOG_BYTES))

    logError('main', 'after rotation')

    expect(await readLog()).toContain('after rotation')
    expect((await readLog(PREVIOUS_LOG_FILE)).length).toBe(MAX_LOG_BYTES)
  })

  it('a folder that cannot be written does not throw', async () => {
    await writeFile(join(directory, LOG_FILE), '')
    useLogFolder(join(directory, LOG_FILE, 'not-a-folder'))

    expect(() => logError('main', 'lost')).not.toThrow()
  })
})
