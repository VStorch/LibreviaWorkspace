import { z } from 'zod'
import { IpcChannel } from './ipc-channels.js'
import type { SerializedError } from './errors.js'

/**
 * Contratos de IPC.
 *
 * Princípio de arquitetura (ver docs/00-plano-tecnico.md §4.5): esta fronteira é
 * **neutra em linguagem**. Só trafegam dados serializáveis — nunca objetos do
 * Node, Buffers compartilhados, classes ou tipos de bibliotecas de terceiros.
 * É isso que permite, no futuro, trocar um worker Node por um binário de outra
 * linguagem sem tocar no resto do aplicativo.
 */

export const appPingRequestSchema = z.object({
  message: z.string().min(1).max(200),
})

export const appPingResponseSchema = z.object({
  echo: z.string(),
  receivedAt: z.number().int(),
  versions: z.object({
    app: z.string(),
    electron: z.string(),
    chrome: z.string(),
    node: z.string(),
  }),
})

export const ipcContracts = {
  [IpcChannel.AppPing]: {
    request: appPingRequestSchema,
    response: appPingResponseSchema,
  },
} as const

export type IpcContracts = typeof ipcContracts

export type IpcRequest<C extends IpcChannel> = z.infer<IpcContracts[C]['request']>
export type IpcResponse<C extends IpcChannel> = z.infer<IpcContracts[C]['response']>

/**
 * Envelope de resultado. Handlers nunca propagam exceções pelo IPC: toda
 * chamada devolve sucesso ou um erro já higienizado.
 */
export type IpcResult<T> =
  { readonly ok: true; readonly data: T } | { readonly ok: false; readonly error: SerializedError }
