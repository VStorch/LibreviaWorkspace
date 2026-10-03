/** O sidecar pode morrer a qualquer momento, e isso não pode custar o documento aberto. */

import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { AppError, ErrorCode } from '@shared/errors.js'
import {
  EMPTY_BINARY,
  FrameReader,
  SidecarMethod,
  encodeFrame,
  healthResultSchema,
  parseResponse,
  type HealthResult,
} from './protocol.js'
import { t } from '../i18n.js'

/** Passou disso, algo travou, e travar em silêncio é pior que falhar rápido. */
export const REQUEST_TIMEOUT_MS = 60_000
export const HEALTH_TIMEOUT_MS = 10_000

export const SHUTDOWN_GRACE_MS = 2_000

const died = (): string => t('errors.sidecar.died')
const timedOut = (): string => t('errors.sidecar.timedOut')

export interface SidecarReply {
  readonly result: unknown
  readonly binary: Uint8Array
}

interface Pending {
  readonly resolve: (reply: SidecarReply) => void
  readonly reject: (error: AppError) => void
  readonly timer: NodeJS.Timeout
}

/** Por parâmetro, para testar contra um sidecar de mentira. */
export type ResolveExecutable = () => Promise<string>

export class SidecarClient {
  readonly #resolveExecutable: ResolveExecutable
  #child: ChildProcessWithoutNullStreams | null = null
  #starting: Promise<ChildProcessWithoutNullStreams> | null = null
  #reader = new FrameReader()
  #pending = new Map<number, Pending>()
  #nextId = 1
  #disposed = false

  constructor(resolveExecutable: ResolveExecutable) {
    this.#resolveExecutable = resolveExecutable
  }

  async health(): Promise<HealthResult> {
    const { result } = await this.request(SidecarMethod.Health, {}, EMPTY_BINARY, HEALTH_TIMEOUT_MS)

    const parsed = healthResultSchema.safeParse(result)
    if (!parsed.success) {
      throw new AppError(ErrorCode.SidecarFailed, died(), t('errors.sidecar.healthContract'))
    }
    return parsed.data
  }

  async request(
    method: SidecarMethod,
    params: unknown,
    binary: Uint8Array = EMPTY_BINARY,
    timeoutMs: number = REQUEST_TIMEOUT_MS,
  ): Promise<SidecarReply> {
    if (this.#disposed) {
      throw new AppError(ErrorCode.SidecarUnavailable, died(), t('errors.sidecar.alreadyClosed'))
    }

    const child = await this.#ensureStarted()

    // `dispose()` pode acontecer enquanto o processo sobe: o pedido ficaria
    // pendurado e o processo, órfão.
    if (this.#disposed) {
      this.#kill()
      throw new AppError(ErrorCode.SidecarUnavailable, died(), t('errors.sidecar.closedDuringRequest'))
    }

    const id = this.#nextId++

    return new Promise<SidecarReply>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(id)
        // Pode estar num laço infinito: derrubar garante que o próximo comece limpo.
        this.#kill()
        reject(new AppError(ErrorCode.SidecarTimeout, timedOut()))
      }, timeoutMs)
      timer.unref()

      this.#pending.set(id, { resolve, reject, timer })

      try {
        child.stdin.write(encodeFrame({ id, method, params }, binary))
      } catch (cause) {
        this.#settle(id, (pending) =>
          pending.reject(
            new AppError(ErrorCode.SidecarFailed, died(), cause instanceof Error ? cause.message : undefined),
          ),
        )
      }
    })
  }

/** Idempotente. */
  dispose(): void {
    this.#disposed = true
    this.#failAllPending(new AppError(ErrorCode.SidecarUnavailable, died(), 'aplicativo encerrando'))

    const child = this.#child
    if (child === null) return
    this.#child = null

    child.stdin.end()
    child.kill('SIGTERM')

    // Se ignorar o SIGTERM, não fica pendurado segurando o encerramento do app.
    const forceKill = setTimeout(() => child.kill('SIGKILL'), SHUTDOWN_GRACE_MS)
    forceKill.unref()
    child.once('exit', () => clearTimeout(forceKill))
  }


  async #ensureStarted(): Promise<ChildProcessWithoutNullStreams> {
    if (this.#child !== null) return this.#child
    // Dois pedidos simultâneos com o processo caído não podem subir dois.
    this.#starting ??= this.#start().finally(() => {
      this.#starting = null
    })
    return this.#starting
  }

  async #start(): Promise<ChildProcessWithoutNullStreams> {
    const executable = await this.#resolveExecutable()

    const child = spawn(executable, [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      shell: false,
      windowsHide: true,
    })

    child.stdout.on('data', (chunk: Buffer) => this.#onStdout(chunk))
    // Diagnóstico nosso: fica no log, e nunca chega ao usuário.
    child.stderr.on('data', (chunk: Buffer) => {
      console.error('[sidecar]', chunk.toString('utf8').trimEnd())
    })

    child.once('error', (cause) => {
      this.#child = null
      this.#failAllPending(new AppError(ErrorCode.SidecarUnavailable, died(), cause.message))
    })

    child.once('exit', (code, signal) => {
      this.#child = null
      this.#reader = new FrameReader()
      this.#failAllPending(
        new AppError(
          ErrorCode.SidecarFailed,
          died(),
          t('errors.sidecar.exitCodeSignal', { code: String(code), signal: String(signal) }),
        ),
      )
    })

    this.#child = child
    return child
  }

  #onStdout(chunk: Buffer): void {
    let frames
    try {
      frames = this.#reader.push(new Uint8Array(chunk))
    } catch (cause) {
      // Fluxo corrompido: não dá para saber onde o próximo quadro começa.
      this.#kill()
      this.#failAllPending(
        cause instanceof AppError ? cause : new AppError(ErrorCode.SidecarFailed, died()),
      )
      return
    }

    for (const frame of frames) {
      let response
      try {
        response = parseResponse(frame.json)
      } catch (cause) {
        this.#failAllPending(cause instanceof AppError ? cause : new AppError(ErrorCode.SidecarFailed, died()))
        return
      }

      this.#settle(response.id, (pending) => {
        if (response.ok) {
          pending.resolve({ result: response.result, binary: frame.binary })
        } else {
          // A frase já vem em português; o código vira SidecarFailed, com o detalhe no log.
          pending.reject(
            new AppError(ErrorCode.SidecarFailed, response.error.message, response.error.code),
          )
        }
      })
    }
  }

  #settle(id: number, apply: (pending: Pending) => void): void {
    const pending = this.#pending.get(id)
    // Já expirou, ou o sidecar inventou um id: não há a quem entregar.
    if (pending === undefined) return

    this.#pending.delete(id)
    clearTimeout(pending.timer)
    apply(pending)
  }

  #failAllPending(error: AppError): void {
    const pendings = [...this.#pending.values()]
    this.#pending.clear()
    for (const pending of pendings) {
      clearTimeout(pending.timer)
      pending.reject(error)
    }
  }

  #kill(): void {
    const child = this.#child
    if (child === null) return
    this.#child = null
    child.kill('SIGKILL')
  }
}
