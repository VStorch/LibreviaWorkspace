/**
 * Cliente do sidecar .NET: sobe o processo, conversa por stdio e garante que
 * ele nunca trave o aplicativo.
 *
 * A regra que orienta todo este arquivo: **o sidecar pode morrer a qualquer
 * momento, e isso não pode custar o documento aberto.** Ele não grava nada em
 * disco e não guarda estado que só exista nele; se cair, o pior caso é a
 * operação em curso falhar com uma frase compreensível, e a próxima subir um
 * processo novo.
 */

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

/**
 * Teto por operação. Documento grande demora, mas nada aqui justifica um
 * minuto: passou disso, alguma coisa travou, e travar em silêncio é pior que
 * falhar rápido.
 */
export const REQUEST_TIMEOUT_MS = 60_000
export const HEALTH_TIMEOUT_MS = 10_000

/** Prazo entre pedir para encerrar e matar de vez. */
export const SHUTDOWN_GRACE_MS = 2_000

const DIED = 'O serviço de formatos foi encerrado inesperadamente. Seu documento continua aberto e intacto.'
const TIMED_OUT =
  'O serviço de formatos demorou demais para responder e a operação foi cancelada. Seu documento continua aberto e intacto.'

export interface SidecarReply {
  readonly result: unknown
  readonly binary: Uint8Array
}

interface Pending {
  readonly resolve: (reply: SidecarReply) => void
  readonly reject: (error: AppError) => void
  readonly timer: NodeJS.Timeout
}

/**
 * Como achar o executável. Recebido por parâmetro para que o cliente não
 * dependa do Electron e possa ser testado contra um sidecar de mentira.
 */
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
      throw new AppError(ErrorCode.SidecarFailed, DIED, 'health fora do contrato')
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
      throw new AppError(ErrorCode.SidecarUnavailable, DIED, 'cliente já encerrado')
    }

    const child = await this.#ensureStarted()

    // Subir o processo leva tempo, e `dispose()` pode acontecer no meio disso —
    // é o caso de fechar o aplicativo logo depois de mandar salvar. Sem esta
    // segunda checagem o pedido se registraria num cliente já encerrado e
    // ficaria pendurado para sempre, e o processo recém-nascido viraria órfão.
    if (this.#disposed) {
      this.#kill()
      throw new AppError(ErrorCode.SidecarUnavailable, DIED, 'encerrado durante o pedido')
    }

    const id = this.#nextId++

    return new Promise<SidecarReply>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(id)
        // Um processo que estourou o prazo não volta ao normal sozinho: pode
        // estar num laço infinito com um documento malformado. Derrubar é o
        // único jeito de garantir que o próximo pedido comece limpo.
        this.#kill()
        reject(new AppError(ErrorCode.SidecarTimeout, TIMED_OUT))
      }, timeoutMs)
      timer.unref()

      this.#pending.set(id, { resolve, reject, timer })

      try {
        child.stdin.write(encodeFrame({ id, method, params }, binary))
      } catch (cause) {
        this.#settle(id, (pending) =>
          pending.reject(
            new AppError(ErrorCode.SidecarFailed, DIED, cause instanceof Error ? cause.message : undefined),
          ),
        )
      }
    })
  }

  /** Encerra sem deixar processo órfão. Idempotente. */
  dispose(): void {
    this.#disposed = true
    this.#failAllPending(new AppError(ErrorCode.SidecarUnavailable, DIED, 'aplicativo encerrando'))

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

  // --- interno -------------------------------------------------------------

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
      // Sem shell: o caminho não passa por interpretação nenhuma.
      shell: false,
      windowsHide: true,
    })

    child.stdout.on('data', (chunk: Buffer) => this.#onStdout(chunk))
    // stderr é diagnóstico nosso e fica no log do main — nunca chega ao usuário.
    child.stderr.on('data', (chunk: Buffer) => {
      console.error('[sidecar]', chunk.toString('utf8').trimEnd())
    })

    child.once('error', (cause) => {
      this.#child = null
      this.#failAllPending(new AppError(ErrorCode.SidecarUnavailable, DIED, cause.message))
    })

    child.once('exit', (code, signal) => {
      this.#child = null
      this.#reader = new FrameReader()
      this.#failAllPending(new AppError(ErrorCode.SidecarFailed, DIED, `saiu com code=${code} signal=${signal}`))
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
        cause instanceof AppError ? cause : new AppError(ErrorCode.SidecarFailed, DIED),
      )
      return
    }

    for (const frame of frames) {
      let response
      try {
        response = parseResponse(frame.json)
      } catch (cause) {
        this.#failAllPending(cause instanceof AppError ? cause : new AppError(ErrorCode.SidecarFailed, DIED))
        return
      }

      this.#settle(response.id, (pending) => {
        if (response.ok) {
          pending.resolve({ result: response.result, binary: frame.binary })
        } else {
          // O sidecar já manda a frase em português; o código dele é interno,
          // então vira SidecarFailed com o detalhe preservado para o log.
          pending.reject(
            new AppError(ErrorCode.SidecarFailed, response.error.message, response.error.code),
          )
        }
      })
    }
  }

  #settle(id: number, apply: (pending: Pending) => void): void {
    const pending = this.#pending.get(id)
    // Resposta sem pedido correspondente: já expirou, ou o sidecar inventou um
    // id. Ignorar é o certo — não há a quem entregar.
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
