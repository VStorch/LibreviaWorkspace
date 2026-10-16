import type { ChildProcessWithoutNullStreams } from 'node:child_process'
import type { Readable } from 'node:stream'
import { transcriptPrompt } from '@services/ai/chat.js'
import { parseClaudeCliLine, type ClaudeCliEvent } from '@services/ai/claude-stream.js'
import type { AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import {
  collect,
  exitOf,
  makeWorkFolder,
  removeWorkFolder,
  spawnWithoutShell,
  terminate,
  type Spawn,
  type WorkFolder,
} from './child-process.js'
import { aiError, refusedError } from './errors.js'
import { takeBoundedLines } from './stream-lines.js'

type CliResult = Extract<ClaudeCliEvent, { type: 'result' }>

export interface ClaudeCliOptions extends WorkFolder {
  readonly command: string
  readonly spawn?: Spawn
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
  const cwd = await makeWorkFolder(options)
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
    await removeWorkFolder(options, cwd)
  }
}

function startCli(options: ClaudeCliOptions, prompt: AiPrompt, cwd: string): RunningCli {
  const child = (options.spawn ?? spawnWithoutShell)(options.command, claudeCliArgs(prompt), { cwd })
  const running = { child, exited: exitOf(child, cliMissing), stderr: collect(child.stderr) }
  child.stdin.on('error', () => undefined)
  child.stdin.end(transcriptPrompt(prompt.messages))
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

function cliMissing(cause: Error): Error {
  return aiError(ErrorCode.AiNotConfigured, 'errors.ai.cliMissing', cause.message)
}
