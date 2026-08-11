import { z } from 'zod'
import { IpcChannel, INVOCABLE_IPC_CHANNELS, type InvocableIpcChannel } from './ipc-channels.js'
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

/**
 * Teto de conteúdo em memória.
 *
 * Um `.sdoc` carrega imagens embutidas como data URI, então é bem maior que um
 * `.txt` equivalente. O limite existe para que um arquivo absurdo não trave a
 * interface, não para restringir uso legítimo.
 */
export const MAX_TEXT_LENGTH = 50_000_000

const documentKindSchema = z.enum(['document', 'spreadsheet'])

const loadedFileSchema = z.object({
  path: z.string(),
  name: z.string(),
  kind: documentKindSchema,
  content: z.string(),
})

const recentFileSchema = z.object({
  path: z.string(),
  name: z.string(),
  kind: documentKindSchema,
  openedAt: z.number().int(),
})

const emptyRequest = z.object({})

/** Diálogo cancelado não é erro: é um desfecho previsto. */
const openResultSchema = z.discriminatedUnion('canceled', [
  z.object({ canceled: z.literal(true) }),
  z.object({ canceled: z.literal(false), file: loadedFileSchema }),
])

const saveResultSchema = z.discriminatedUnion('canceled', [
  z.object({ canceled: z.literal(true) }),
  z.object({ canceled: z.literal(false), path: z.string(), name: z.string() }),
])

export const ipcContracts = {
  [IpcChannel.FileOpen]: {
    request: emptyRequest,
    response: openResultSchema,
  },
  [IpcChannel.FileOpenRecent]: {
    request: z.object({ path: z.string().min(1) }),
    response: z.object({ file: loadedFileSchema }),
  },
  [IpcChannel.FileSave]: {
    request: z.object({
      path: z.string().min(1),
      content: z.string().max(MAX_TEXT_LENGTH),
    }),
    response: z.object({ path: z.string(), name: z.string() }),
  },
  [IpcChannel.FileChooseSavePath]: {
    request: z.object({ suggestedName: z.string().min(1).max(255) }),
    response: saveResultSchema,
  },
  [IpcChannel.RecentList]: {
    request: emptyRequest,
    response: z.object({ files: z.array(recentFileSchema) }),
  },
  [IpcChannel.RecentClear]: {
    request: emptyRequest,
    response: z.object({ files: z.array(recentFileSchema) }),
  },
  [IpcChannel.ImagePick]: {
    request: emptyRequest,
    response: z.discriminatedUnion('canceled', [
      z.object({ canceled: z.literal(true) }),
      z.object({
        canceled: z.literal(false),
        // Data URI já validado por assinatura de bytes no processo main.
        dataUrl: z.string(),
        name: z.string(),
      }),
    ]),
  },
  [IpcChannel.DialogConfirmDiscard]: {
    request: z.object({ fileName: z.string().min(1).max(255) }),
    response: z.object({ choice: z.enum(['save', 'discard', 'cancel']) }),
  },
  [IpcChannel.DialogConfirmPlainText]: {
    request: z.object({ fileName: z.string().min(1).max(255) }),
    response: z.object({ choice: z.enum(['keep-plain', 'save-as-document', 'cancel']) }),
  },
  [IpcChannel.WindowSetState]: {
    request: z.object({
      title: z.string().max(300),
      isDirty: z.boolean(),
    }),
    response: z.object({ applied: z.literal(true) }),
  },
  [IpcChannel.WindowClose]: {
    request: emptyRequest,
    response: z.object({ closing: z.literal(true) }),
  },
} as const

export type IpcContracts = typeof ipcContracts

export type IpcRequest<C extends InvocableIpcChannel> = z.infer<IpcContracts[C]['request']>
export type IpcResponse<C extends InvocableIpcChannel> = z.infer<IpcContracts[C]['response']>

/**
 * Envelope de resultado. Handlers nunca propagam exceções pelo IPC: toda
 * chamada devolve sucesso ou um erro já higienizado.
 */
export type IpcResult<T> =
  { readonly ok: true; readonly data: T } | { readonly ok: false; readonly error: SerializedError }

export { INVOCABLE_IPC_CHANNELS }
