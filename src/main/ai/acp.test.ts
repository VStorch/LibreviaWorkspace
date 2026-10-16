import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'
import type { ChildProcessWithoutNullStreams } from 'node:child_process'
import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { createAcpProvider } from './acp.js'

const PROMPT = {
  model: '',
  system: 'regras',
  messages: [{ role: 'user' as const, text: 'Olá' }],
  maxTokens: 100,
}

type Message = { id?: number | string; method?: string; params?: Record<string, unknown>; result?: unknown }

interface FakeAgentBehavior {
  readonly authRequired?: boolean
  readonly ignoresCancel?: boolean
  readonly stopReason?: string
  readonly missing?: boolean
  readonly notAgent?: 'methodNotFound' | 'closes'
}

const MODEL_OPTION = {
  id: 'model',
  name: 'Modelo',
  type: 'select',
  category: 'model',
  currentValue: 'rapido',
  options: [
    { value: 'rapido', name: 'Rápido' },
    { value: 'esperto', name: 'Esperto' },
  ],
}

/** Speaks ACP by hand, line by line, the way a real agent process would on its stdio. */
class FakeAgent extends EventEmitter {
  readonly stdin = new PassThrough()
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  exitCode: number | null = null
  signalCode: string | null = null
  readonly received: Message[] = []
  readonly signals: string[] = []
  private buffer = ''
  private promptId: number | string | undefined

  constructor(private readonly behavior: FakeAgentBehavior) {
    super()
    this.stdin.on('data', (chunk: Buffer) => this.read(chunk.toString()))
  }

  kill(signal: string): boolean {
    this.signals.push(signal)
    if (this.exitCode !== null || this.signalCode !== null) return true
    this.signalCode = signal
    this.stdout.end()
    this.stderr.end()
    queueMicrotask(() => {
      this.emit('exit', null, signal)
      this.emit('close', null, signal)
    })
    return true
  }

  private closeAlone(): void {
    this.exitCode = 1
    this.stdout.end()
    this.stderr.end()
    queueMicrotask(() => {
      this.emit('exit', 1, null)
      this.emit('close', 1, null)
    })
  }

  private send(message: Record<string, unknown>): void {
    if (!this.stdout.writableEnded) this.stdout.write(`${JSON.stringify({ jsonrpc: '2.0', ...message })}\n`)
  }

  private update(update: Record<string, unknown>): void {
    this.send({ method: 'session/update', params: { sessionId: 's1', update } })
  }

  private read(text: string): void {
    this.buffer += text
    let newline = this.buffer.indexOf('\n')
    while (newline >= 0) {
      const message = JSON.parse(this.buffer.slice(0, newline)) as Message
      this.buffer = this.buffer.slice(newline + 1)
      this.received.push(message)
      this.answer(message)
      newline = this.buffer.indexOf('\n')
    }
  }

  private answer(message: Message): void {
    if (this.behavior.missing) return
    switch (message.method) {
      case 'initialize':
        if (this.behavior.notAgent === 'methodNotFound')
          return this.send({ id: message.id, error: { code: -32601, message: 'Method not found' } })
        if (this.behavior.notAgent === 'closes') return this.closeAlone()
        if (this.behavior.authRequired)
          return this.send({ id: message.id, error: { code: -32000, message: 'Authentication required' } })
        return this.send({ id: message.id, result: { protocolVersion: 1, agentCapabilities: {} } })
      case 'session/new':
        return this.send({ id: message.id, result: { sessionId: 's1', configOptions: [MODEL_OPTION] } })
      case 'session/set_config_option':
        return this.send({ id: message.id, result: { configOptions: [MODEL_OPTION] } })
      case 'session/prompt':
        return this.prompt(message.id!)
      case 'session/cancel':
        if (!this.behavior.ignoresCancel)
          this.send({ id: this.promptId, result: { stopReason: 'cancelled' } })
        return
      default:
        if (message.id === 'p1') this.finishPrompt()
    }
  }

  private prompt(id: number | string): void {
    this.promptId = id
    this.update({ sessionUpdate: 'agent_thought_chunk', content: { type: 'text', text: 'pensando' } })
    this.update({ sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'Olá, ' } })
    this.update({ sessionUpdate: 'tool_call', toolCallId: 't1', title: 'Ler arquivo' })
    if (this.behavior.ignoresCancel || this.behavior.stopReason === 'cancelled') return
    this.send({
      id: 'p1',
      method: 'session/request_permission',
      params: {
        sessionId: 's1',
        toolCall: { toolCallId: 't1' },
        options: [
          { optionId: 'sim', name: 'Permitir', kind: 'allow_once' },
          { optionId: 'nao', name: 'Recusar', kind: 'reject_once' },
        ],
      },
    })
  }

  private finishPrompt(): void {
    this.update({ sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'mundo' } })
    this.send({ id: this.promptId, result: { stopReason: this.behavior.stopReason ?? 'end_turn' } })
  }
}

function setup(behavior: FakeAgentBehavior = {}) {
  const child = new FakeAgent(behavior)
  const calls: { command: string; args: readonly string[]; cwd: string }[] = []
  const removed: string[] = []
  const provider = createAcpProvider({
    command: '/usr/bin/agente',
    args: ['--acp'],
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

const methodsOf = (child: FakeAgent): (string | undefined)[] =>
  child.received.map((message) => message.method)

describe('createAcpProvider', () => {
  it('abre o agente sem shell na pasta vazia, sem arquivos nem terminal, e repassa só o texto da resposta', async () => {
    const { child, calls, removed, provider } = setup()
    const parts = await collect(provider.stream(PROMPT, new AbortController().signal))

    expect(calls).toEqual([{ command: '/usr/bin/agente', args: ['--acp'], cwd: '/tmp/librevia-ai-x' }])
    expect(parts).toEqual([{ model: 'rapido' }, { text: 'Olá, ' }, { text: 'mundo' }])
    const initialize = child.received.find((message) => message.method === 'initialize')
    expect(initialize?.params?.['clientCapabilities']).toEqual({
      fs: { readTextFile: false, writeTextFile: false },
      terminal: false,
    })
    const session = child.received.find((message) => message.method === 'session/new')
    expect(session?.params).toEqual({ cwd: '/tmp/librevia-ai-x', mcpServers: [] })
    const prompt = child.received.find((message) => message.method === 'session/prompt')
    expect(prompt?.params?.['prompt']).toEqual([{ type: 'text', text: 'regras\n\nOlá' }])
    expect(child.signals).toContain('SIGTERM')
    expect(removed).toEqual(['/tmp/librevia-ai-x'])
  })

  it('recusa o pedido de permissão do agente', async () => {
    const { child, provider } = setup()
    await collect(provider.stream(PROMPT, new AbortController().signal))
    const answer = child.received.find((message) => message.id === 'p1')
    expect(answer?.result).toEqual({ outcome: { outcome: 'selected', optionId: 'nao' } })
  })

  it('o modelo escolhido vai para a sessão antes do pedido', async () => {
    const { child, provider } = setup()
    const parts = await collect(
      provider.stream({ ...PROMPT, model: 'esperto' }, new AbortController().signal),
    )
    const chosen = child.received.find((message) => message.method === 'session/set_config_option')
    expect(chosen?.params).toEqual({ sessionId: 's1', configId: 'model', value: 'esperto' })
    expect(methodsOf(child).indexOf('session/set_config_option')).toBeLessThan(
      methodsOf(child).indexOf('session/prompt'),
    )
    expect(parts[0]).toEqual({ model: 'esperto' })
  })

  it('lista os modelos da sessão e encerra o processo', async () => {
    const { child, removed, provider } = setup()
    const models = await provider.listModels(new AbortController().signal)
    expect(models).toEqual([
      { id: 'rapido', label: 'Rápido' },
      { id: 'esperto', label: 'Esperto' },
    ])
    expect(methodsOf(child)).not.toContain('session/prompt')
    expect(child.signals).toContain('SIGTERM')
    expect(removed).toEqual(['/tmp/librevia-ai-x'])
  })

  it('agente que pede login vira aviso para fazer login no terminal', async () => {
    const { provider } = setup({ authRequired: true })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiAuthFailed,
    })
  })

  it('recusa do modelo vira o erro de recusa', async () => {
    const { provider } = setup({ stopReason: 'refusal' })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiRefused,
    })
  })

  it('cancelar avisa o agente e encerra o processo', async () => {
    const { child, removed, provider } = setup({ stopReason: 'cancelled' })
    const controller = new AbortController()
    const parts: unknown[] = []
    for await (const part of provider.stream(PROMPT, controller.signal)) {
      parts.push(part)
      if ('text' in (part as object)) controller.abort()
    }
    expect(methodsOf(child)).toContain('session/cancel')
    expect(child.signals).toContain('SIGTERM')
    expect(removed).toEqual(['/tmp/librevia-ai-x'])
  })

  it('agente que ignora o cancelamento é encerrado mesmo assim', async () => {
    const { child, provider } = setup({ ignoresCancel: true })
    const controller = new AbortController()
    const run = collect(
      (async function* () {
        for await (const part of provider.stream(PROMPT, controller.signal)) {
          if ('text' in (part as object)) controller.abort()
          yield part
        }
      })(),
    )
    await expect(run).rejects.toBeDefined()
    expect(methodsOf(child)).toContain('session/cancel')
    expect(child.signals).toContain('SIGTERM')
  })

  it.each(['methodNotFound', 'closes'] as const)(
    'programa que não fala ACP (%s) vira aviso legível, sem texto técnico',
    async (notAgent) => {
      const { provider } = setup({ notAgent })
      const failure = await collect(provider.stream(PROMPT, new AbortController().signal)).catch(
        (cause: unknown) => cause,
      )
      expect(failure).toMatchObject({ code: ErrorCode.AiFailed, detail: undefined })
      expect(String((failure as Error).message)).toMatch(/ACP/)
      expect(String((failure as Error).message)).not.toMatch(/Method not found|connection closed/i)
    },
  )

  it('programa ausente vira aviso de agente não encontrado', async () => {
    const { child, provider } = setup({ missing: true })
    const run = collect(provider.stream(PROMPT, new AbortController().signal))
    setTimeout(() => child.emit('error', new Error('spawn ENOENT')), 0)
    await expect(run).rejects.toMatchObject({ code: ErrorCode.AiNotConfigured })
  })
})
