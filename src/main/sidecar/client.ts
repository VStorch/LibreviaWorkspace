/** The sidecar may die at any moment, and that must not cost the open document. */

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

/** Past this, something is stuck, and hanging silently is worse than failing fast. */
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

/** A parameter, to test against a fake sidecar. */
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

    // `dispose()` may happen while the process starts: the request would hang and the process would
    // be orphaned.
    if (this.#disposed) {
      this.#kill()
      throw new AppError(ErrorCode.SidecarUnavailable, died(), t('errors.sidecar.closedDuringRequest'))
    }

    const id = this.#nextId++

    return new Promise<SidecarReply>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(id)
        // It may be in an infinite loop: killing it guarantees the next one starts clean.
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

  dispose(): void {
    this.#disposed = true
    this.#failAllPending(new AppError(ErrorCode.SidecarUnavailable, died(), 'aplicativo encerrando'))

    const child = this.#child
    if (child === null) return
    this.#child = null

    child.stdin.end()
    child.kill('SIGTERM')

    // If it ignores SIGTERM, it does not hang around holding up the app shutdown.
    const forceKill = setTimeout(() => child.kill('SIGKILL'), SHUTDOWN_GRACE_MS)
    forceKill.unref()
    child.once('exit', () => clearTimeout(forceKill))
  }


  async #ensureStarted(): Promise<ChildProcessWithoutNullStreams> {
    if (this.#child !== null) return this.#child
    // Two simultaneous requests with the process down must not start two.
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
    // Our own diagnostics: they stay in the log and never reach the user.
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
      // Corrupt stream: there is no way to know where the next frame starts.
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
          // The sentence is already Portuguese; the code becomes SidecarFailed, with the detail in
          // the log.
          pending.reject(
            new AppError(ErrorCode.SidecarFailed, response.error.message, response.error.code),
          )
        }
      })
    }
  }

  #settle(id: number, apply: (pending: Pending) => void): void {
    const pending = this.#pending.get(id)
    // Already timed out, or the sidecar made up an id: nobody to deliver to.
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
