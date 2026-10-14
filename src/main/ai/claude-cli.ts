import { spawn as nodeSpawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Readable } from 'node:stream'
import { parseClaudeCliLine, type ClaudeCliEvent } from '@services/ai/claude-stream.js'
import type { AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import { aiError, refusedError } from './errors.js'
import { takeBoundedLines } from './stream-lines.js'

const KILL_GRACE_MS = 2_000
const MAX_STDERR_LENGTH = 2_000

type CliResult = Extract<ClaudeCliEvent, { type: 'result' }>
type Spawn = (
  command: string,
  args: readonly string[],
  options: { cwd: string },
) => ChildProcessWithoutNullStreams

export interface ClaudeCliOptions {
  readonly command: string
  readonly spawn?: Spawn
  readonly makeTempDir?: () => Promise<string>
  readonly removeDir?: (dir: string) => Promise<void>
}

/**
 * No tools, no MCP servers, no settings or session files: the CLI only turns a prompt into text.
 * The empty working folder keeps it from reading the person's project.
 */
export function claudeCliArgs(prompt: AiPrompt): string[] {
  return [
    '-p',
    '--tools',
    '',
    '--output-format',
    'stream-json',
    '--include-partial-messages',
    '--no-session-persistence',
    '--verbose',
    '--strict-mcp-config',
    '--setting-sources',
    '',
    '--system-prompt',
    prompt.system,
    ...(prompt.model === '' ? [] : ['--model', prompt.model]),
  ]
}

const defaultSpawn: Spawn = (command, args, options) =>
  nodeSpawn(command, args, {
    cwd: options.cwd,
    shell: false,
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  })

export function createClaudeCliProvider(options: ClaudeCliOptions): AiProvider {
  return {
    listModels: () => Promise.resolve([]),
    stream: (prompt, signal) => runCli(options, prompt, signal),
  }
}

interface RunningCli {
  readonly child: ChildProcessWithoutNullStreams
  readonly exited: Promise<number | null>
  readonly stderr: () => string
}

async function* runCli(
  options: ClaudeCliOptions,
  prompt: AiPrompt,
  signal: AbortSignal,
): AsyncGenerator<AiStreamPart> {
  signal.throwIfAborted()
  const cwd = await (options.makeTempDir ?? (() => mkdtemp(join(tmpdir(), 'librevia-ai-'))))()
  let running: RunningCli | undefined
  const stop = (): void => {
    if (running !== undefined) terminate(running.child)
  }
  try {
    running = startCli(options, prompt, cwd)
    signal.addEventListener('abort', stop, { once: true })
    if (signal.aborted) stop()
    yield* readCli(running, signal)
  } finally {
    signal.removeEventListener('abort', stop)
    stop()
    void running?.exited.catch(() => undefined)
    await (options.removeDir ?? ((dir) => rm(dir, { recursive: true, force: true })))(cwd)
  }
}

function startCli(options: ClaudeCliOptions, prompt: AiPrompt, cwd: string): RunningCli {
  const child = (options.spawn ?? defaultSpawn)(options.command, claudeCliArgs(prompt), { cwd })
  const running = { child, exited: exitOf(child), stderr: collect(child.stderr) }
  child.stdin.on('error', () => undefined)
  child.stdin.end(prompt.user)
  return running
}

async function* readCli(running: RunningCli, signal: AbortSignal): AsyncGenerator<AiStreamPart> {
  let result: CliResult | undefined
  for await (const event of cliEvents(running.child.stdout)) {
    if (event.type === 'result') result = event
    else if (event.type === 'text') yield { text: event.text }
    else if (event.type === 'model') yield { model: event.model }
  }
  const code = await running.exited
  signal.throwIfAborted()
  yield* finish(result, code, running.stderr())
}

function* finish(
  result: CliResult | undefined,
  code: number | null,
  stderr: string,
): Generator<AiStreamPart> {
  if (result === undefined) {
    if (code === 0) return
    throw aiError(ErrorCode.AiFailed, 'errors.ai.failed', stderr.trim())
  }
  if (result.refused) throw refusedError()
  if (result.failed) throw aiError(ErrorCode.AiFailed, 'errors.ai.failed', result.message)
  if (result.costUsd !== undefined) yield { costUsd: result.costUsd }
}

async function* cliEvents(stdout: Readable): AsyncGenerator<ClaudeCliEvent> {
  const decoder = new TextDecoder()
  let buffer = ''
  for await (const chunk of stdout) {
    const split = takeBoundedLines(buffer + decoder.decode(chunk as Uint8Array, { stream: true }))
    buffer = split.rest
    for (const line of split.lines) yield parseClaudeCliLine(line)
  }
  if (buffer.trim() !== '') yield parseClaudeCliLine(buffer)
}

function exitOf(child: ChildProcessWithoutNullStreams): Promise<number | null> {
  return new Promise((resolve, reject) => {
    child.once('error', (cause) =>
      reject(aiError(ErrorCode.AiNotConfigured, 'errors.ai.cliMissing', cause.message)),
    )
    child.once('close', (code) => resolve(code))
  })
}

function collect(stream: Readable): () => string {
  let text = ''
  stream.on('data', (chunk: Buffer | string) => {
    text = (text + chunk.toString()).slice(-MAX_STDERR_LENGTH)
  })
  return () => text
}

function terminate(child: ChildProcessWithoutNullStreams): void {
  if (child.exitCode !== null || child.signalCode !== null) return
  child.kill('SIGTERM')
  const timer = setTimeout(() => child.kill('SIGKILL'), KILL_GRACE_MS)
  timer.unref()
  child.once('exit', () => clearTimeout(timer))
}
