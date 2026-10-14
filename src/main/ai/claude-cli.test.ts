import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'
import type { ChildProcessWithoutNullStreams } from 'node:child_process'
import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { claudeCliArgs, createClaudeCliProvider } from './claude-cli.js'

const PROMPT = { model: '', system: 'regras', user: 'Olá', maxTokens: 100 }

class FakeChild extends EventEmitter {
  readonly stdin = new PassThrough()
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  exitCode: number | null = null
  signalCode: string | null = null
  readonly signals: string[] = []
  stdinText = ''

  constructor() {
    super()
    this.stdin.on('data', (chunk: Buffer) => (this.stdinText += chunk.toString()))
  }

  kill(signal: string): boolean {
    this.signals.push(signal)
    this.finish(null, signal)
    return true
  }

  finish(code: number | null, signal: string | null = null): void {
    if (this.exitCode !== null || this.signalCode !== null) return
    this.exitCode = code
    this.signalCode = signal
    this.stdout.end()
    this.stderr.end()
    queueMicrotask(() => {
      this.emit('exit', code, signal)
      this.emit('close', code, signal)
    })
  }
}

const line = (value: unknown): string => `${JSON.stringify(value)}\n`
const delta = (text: string): string =>
  line({ type: 'stream_event', event: { type: 'content_block_delta', delta: { type: 'text_delta', text } } })

function setup() {
  const child = new FakeChild()
  const calls: { command: string; args: readonly string[]; cwd: string }[] = []
  const removed: string[] = []
  const provider = createClaudeCliProvider({
    command: '/usr/bin/claude',
    spawn: (command, args, options) => {
      calls.push({ command, args, cwd: options.cwd })
      return child as unknown as ChildProcessWithoutNullStreams
    },
    makeTempDir: async () => '/tmp/librevia-ai-x',
    removeDir: async (dir) => {
      removed.push(dir)
    },
  })
  return { child, calls, removed, provider }
}

async function collect(iterable: AsyncIterable<unknown>): Promise<unknown[]> {
  const parts: unknown[] = []
  for await (const part of iterable) parts.push(part)
  return parts
}

describe('claudeCliArgs', () => {
  it('desliga ferramentas, MCP, configurações e sessão', () => {
    const args = claudeCliArgs(PROMPT)
    expect(args).toEqual(expect.arrayContaining(['-p', '--no-session-persistence', '--strict-mcp-config']))
    expect(args[args.indexOf('--tools') + 1]).toBe('')
    expect(args[args.indexOf('--setting-sources') + 1]).toBe('')
    expect(args[args.indexOf('--system-prompt') + 1]).toBe('regras')
    expect(args).not.toContain('--model')
    expect(claudeCliArgs({ ...PROMPT, model: 'sonnet' }).slice(-2)).toEqual(['--model', 'sonnet'])
  })
})

describe('provedor Claude Code', () => {
  it('manda o pedido pelo stdin, lê o texto e o custo, e apaga a pasta temporária', async () => {
    const { child, calls, removed, provider } = setup()
    const parts = collect(provider.stream(PROMPT, new AbortController().signal))
    await new Promise((resolve) => setImmediate(resolve))
    child.stdout.write(line({ type: 'system', subtype: 'init', model: 'claude-opus-5-5' }) + delta('Olá'))
    child.stdout.write(delta(' mundo') + line({ type: 'result', is_error: false, total_cost_usd: 0.01 }))
    child.finish(0)

    expect(await parts).toEqual([
      { model: 'claude-opus-5-5' },
      { text: 'Olá' },
      { text: ' mundo' },
      { costUsd: 0.01 },
    ])
    expect(calls[0]).toMatchObject({ command: '/usr/bin/claude', cwd: '/tmp/librevia-ai-x' })
    expect(child.stdinText).toBe('Olá')
    expect(removed).toEqual(['/tmp/librevia-ai-x'])
  })

  it('erro informado pelo programa vira erro legível', async () => {
    const { child, provider } = setup()
    const parts = collect(provider.stream(PROMPT, new AbortController().signal))
    await new Promise((resolve) => setImmediate(resolve))
    child.stdout.write(line({ type: 'result', is_error: true, result: 'Not logged in' }))
    child.finish(1)
    await expect(parts).rejects.toMatchObject({ code: ErrorCode.AiFailed, detail: 'Not logged in' })
  })

  it('cancelar encerra o processo e limpa a pasta', async () => {
    const { child, removed, provider } = setup()
    const controller = new AbortController()
    const parts = collect(provider.stream(PROMPT, controller.signal))
    await new Promise((resolve) => setImmediate(resolve))
    child.stdout.write(delta('Olá'))
    controller.abort()

    await expect(parts).rejects.toThrow()
    expect(child.signals).toEqual(['SIGTERM'])
    expect(removed).toEqual(['/tmp/librevia-ai-x'])
  })

  it('pedido já cancelado não abre o programa', async () => {
    const { calls, provider } = setup()
    const controller = new AbortController()
    controller.abort()
    await expect(collect(provider.stream(PROMPT, controller.signal))).rejects.toThrow()
    expect(calls).toHaveLength(0)
  })

  it('cancelado enquanto a pasta temporária era criada encerra o processo logo ao abrir', async () => {
    const child = new FakeChild()
    const controller = new AbortController()
    const removed: string[] = []
    const provider = createClaudeCliProvider({
      command: '/usr/bin/claude',
      spawn: () => child as unknown as ChildProcessWithoutNullStreams,
      makeTempDir: async () => {
        controller.abort()
        return '/tmp/librevia-ai-y'
      },
      removeDir: async (dir) => {
        removed.push(dir)
      },
    })
    await expect(collect(provider.stream(PROMPT, controller.signal))).rejects.toThrow()
    expect(child.signals).toEqual(['SIGTERM'])
    expect(removed).toEqual(['/tmp/librevia-ai-y'])
  })

  it('falha ao abrir o programa ainda apaga a pasta temporária', async () => {
    const removed: string[] = []
    const provider = createClaudeCliProvider({
      command: '/usr/bin/claude',
      spawn: () => {
        throw new Error('spawn EACCES')
      },
      makeTempDir: async () => '/tmp/librevia-ai-z',
      removeDir: async (dir) => {
        removed.push(dir)
      },
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toThrow()
    expect(removed).toEqual(['/tmp/librevia-ai-z'])
  })
})
