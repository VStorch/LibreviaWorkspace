import type { ChildProcessWithoutNullStreams } from 'node:child_process'
import { Readable, Writable } from 'node:stream'
import {
  client,
  methods,
  ndJsonStream,
  PROTOCOL_VERSION,
  type ActiveSession,
  type ClientConnection,
} from '@agentclientprotocol/sdk'
import {
  answerPermission,
  isAuthError,
  modelOptionOf,
  modelsOf,
  textOfUpdate,
} from '@services/ai/acp-session.js'
import { transcriptPrompt } from '@services/ai/chat.js'
import type { AiModel, AiPrompt, AiProvider, AiStreamPart, AiToolServer } from '@services/ai/provider.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import { MAX_AI_STREAM_LINE_LENGTH } from '@shared/limits.js'
import {
  collect,
  exitOf,
  makeWorkFolder,
  removeWorkFolder,
  spawnWithoutShell,
  terminate,
  terminateNow,
  type Spawn,
  type WorkFolder,
} from './child-process.js'
import { aiError, refusedError } from './errors.js'

const STARTUP_TIMEOUT_MS = 60_000
const CANCEL_GRACE_MS = 1_000
const CLIENT_NAME = 'Librevia'

export interface AcpOptions extends WorkFolder {
  readonly command: string
  readonly args: readonly string[]
  readonly spawn?: Spawn
  readonly startupTimeoutMs?: number
}

/** One agent process per request: nothing the agent keeps between requests carries over. */
export function createAcpProvider(options: AcpOptions): AiProvider {
  return {
    listModels: (signal) => listModels(options, signal),
    stream: (prompt, signal) => runPrompt(options, prompt, signal),
  }
}

const runningAgents = new Set<AgentProcess>()

/** On quit: the window's cancellation waits for the agent, and the app does not. */
export function stopAllAgents(): void {
  for (const agent of runningAgents) agent.close({ immediately: true })
}

/** Librevia's MCP server as ACP describes a program the agent starts. */
function stdioServer(server: AiToolServer) {
  return {
    name: server.name,
    command: server.command,
    args: [...server.args],
    env: Object.entries(server.env).map(([name, value]) => ({ name, value })),
  }
}

/**
 * A client that offers the agent nothing of its own: no files, no terminal, every permission
 * refused but to run Librevia's tools, when a message brings them. Otherwise the agent can only
 * answer in text.
 */
class AgentProcess {
  readonly exited: Promise<number | null>
  private readonly child: ChildProcessWithoutNullStreams
  private readonly connection: ClientConnection
  private readonly ended: Promise<never>
  private session: ActiveSession | undefined
  private toolServer: AiToolServer | undefined

  constructor(
    private readonly options: AcpOptions,
    private readonly cwd: string,
  ) {
    this.child = (options.spawn ?? spawnWithoutShell)(options.command, options.args, { cwd })
    this.child.stdin.on('error', () => undefined)
    runningAgents.add(this)
    const stderr = collect(this.child.stderr)
    this.exited = exitOf(this.child, agentMissing)
    this.ended = this.exited.then((code) => {
      throw aiError(ErrorCode.AiFailed, 'errors.ai.failed', stderr().trim() || `exit ${String(code)}`)
    })
    this.ended.catch(() => undefined)
    this.connection = client({ name: CLIENT_NAME })
      .onRequest(methods.client.session.requestPermission, ({ params }) =>
        answerPermission(params.toolCall, params.options, this.toolServer),
      )
      .connect(
        ndJsonStream(
          Writable.toWeb(this.child.stdin) as WritableStream<Uint8Array>,
          Readable.toWeb(this.child.stdout) as ReadableStream<Uint8Array>,
          { maxMessageBytes: MAX_AI_STREAM_LINE_LENGTH },
        ),
      )
  }

  async open(toolServer?: AiToolServer): Promise<ActiveSession> {
    this.toolServer = toolServer
    const agent = this.connection.agent
    const startup = this.options.startupTimeoutMs ?? STARTUP_TIMEOUT_MS
    await this.within(
      agent.request(methods.agent.initialize, {
        protocolVersion: PROTOCOL_VERSION,
        clientCapabilities: { fs: { readTextFile: false, writeTextFile: false }, terminal: false },
      }),
      startup,
      handshakeError,
    )
    const builder = agent.buildSession(this.cwd)
    if (toolServer !== undefined) builder.withMcpServer(stdioServer(toolServer))
    this.session = await this.within(builder.start(), startup, handshakeError)
    return this.session
  }

  async chooseModel(session: ActiveSession, model: string): Promise<string> {
    const option = modelOptionOf(session.newSessionResponse.configOptions)
    if (option === undefined) return ''
    if (model === '' || model === option.current) return option.current
    await this.within(
      this.connection.agent.request(methods.agent.session.setConfigOption, {
        sessionId: session.sessionId,
        configId: option.configId,
        value: model,
      }),
    )
    return model
  }

  /** The process may die without answering; then the pending step fails instead of waiting forever. */
  within<T>(
    step: Promise<T>,
    timeoutMs?: number,
    failure: (cause: unknown) => unknown = (cause) => cause,
  ): Promise<T> {
    const fail = (cause: unknown): never => {
      throw failure(cause)
    }
    const racers: Promise<T>[] = [step.catch(fail), this.ended.catch(fail)]
    if (timeoutMs !== undefined) racers.push(deadline(timeoutMs))
    return Promise.race(racers)
  }

  /** Asks the agent to stop and, if it does not, ends the process anyway. */
  cancel(): void {
    const session = this.session
    if (session === undefined) return this.close()
    this.connection.agent
      .notify(methods.agent.session.cancel, { sessionId: session.sessionId })
      .catch(() => undefined)
    setTimeout(() => this.close(), CANCEL_GRACE_MS).unref()
  }

  close(options: { immediately?: boolean } = {}): void {
    runningAgents.delete(this)
    this.session?.dispose()
    this.connection.close()
    if (options.immediately === true) terminateNow(this.child)
    else terminate(this.child)
  }
}

function deadline(ms: number): Promise<never> {
  return new Promise((_, reject) => {
    setTimeout(() => reject(aiError(ErrorCode.AiFailed, 'errors.ai.agentNoAnswer')), ms).unref()
  })
}

function agentMissing(cause: Error): Error {
  return aiError(ErrorCode.AiNotConfigured, 'errors.ai.agentMissing', cause.message)
}

/**
 * A program that closes the connection or does not know the protocol's first methods is not an
 * agent, or was started without the arguments that make it one.
 */
function handshakeError(cause: unknown): unknown {
  if (isAuthError(cause)) return aiError(ErrorCode.AiAuthFailed, 'errors.ai.agentLogin')
  if (cause instanceof AppError && cause.code !== ErrorCode.AiFailed) return cause
  return aiError(ErrorCode.AiFailed, 'errors.ai.notAnAgent')
}

function readableError(cause: unknown): unknown {
  if (cause instanceof AppError) return cause
  if (isAuthError(cause)) return aiError(ErrorCode.AiAuthFailed, 'errors.ai.agentLogin')
  return aiError(
    ErrorCode.AiFailed,
    'errors.ai.failed',
    cause instanceof Error ? cause.message : String(cause),
  )
}

async function* withAgent<T>(
  options: AcpOptions,
  signal: AbortSignal,
  work: (agent: AgentProcess) => AsyncGenerator<T>,
): AsyncGenerator<T> {
  signal.throwIfAborted()
  const cwd = await makeWorkFolder(options)
  let agent: AgentProcess | undefined
  const cancel = (): void => agent?.cancel()
  try {
    agent = new AgentProcess(options, cwd)
    signal.addEventListener('abort', cancel, { once: true })
    yield* work(agent)
  } catch (cause) {
    throw readableError(cause)
  } finally {
    signal.removeEventListener('abort', cancel)
    agent?.close()
    await agent?.exited.catch(() => undefined)
    await removeWorkFolder(options, cwd)
  }
}

function runPrompt(options: AcpOptions, prompt: AiPrompt, signal: AbortSignal): AsyncGenerator<AiStreamPart> {
  return withAgent(options, signal, (agent) => answer(agent, prompt))
}

async function* answer(agent: AgentProcess, prompt: AiPrompt): AsyncGenerator<AiStreamPart> {
  const session = await agent.open(prompt.toolServer)
  const model = await agent.chooseModel(session, prompt.model)
  if (model !== '') yield { model }
  session.prompt(`${prompt.system}\n\n${transcriptPrompt(prompt.messages)}`).catch(() => undefined)
  for (;;) {
    const message = await agent.within(session.nextUpdate())
    if (message.kind === 'stop') {
      if (message.stopReason === 'refusal') throw refusedError()
      return
    }
    const text = textOfUpdate(message.update)
    if (text !== '') yield { text }
  }
}

async function listModels(options: AcpOptions, signal: AbortSignal): Promise<AiModel[]> {
  const models: AiModel[] = []
  const listing = withAgent(options, signal, async function* (agent) {
    const session = await agent.open()
    yield* modelsOf(session.newSessionResponse.configOptions)
  })
  for await (const model of listing) models.push(model)
  return models
}
