import { z } from 'zod'
import { IpcChannel, INVOCABLE_IPC_CHANNELS, type InvocableIpcChannel } from './ipc-channels.js'
import { pageSetupSchema } from './schemas.js'
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

/** Ver `LossInventory` em types.ts: invisível e perdido não são a mesma coisa. */
const inventorySchema = z.object({
  invisible: z.array(z.string().max(300)).max(50),
  lost: z.array(z.string().max(300)).max(50),
  structural: z.array(z.string().max(300)).max(50).default([]),
})

const loadedFileSchema = z.object({
  path: z.string(),
  name: z.string(),
  kind: documentKindSchema,
  content: z.string(),
  inventory: inventorySchema.optional(),
})

const recentFileSchema = z.object({
  path: z.string(),
  name: z.string(),
  kind: documentKindSchema,
  openedAt: z.number().int(),
})

/** O aviso de recuperação mostra de que arquivo veio e de quando é. */
const draftSummarySchema = z.object({
  path: z.string().nullable(),
  name: z.string(),
  kind: documentKindSchema,
  savedAt: z.number().int(),
})

const emptyRequest = z.object({})

/** O HTML vem do editor; a configuração de página, do documento. */
const printRequestSchema = z.object({
  html: z.string().max(MAX_TEXT_LENGTH),
  page: pageSetupSchema,
})

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
    response: z.object({
      path: z.string(),
      name: z.string(),
      inventory: inventorySchema.optional(),
    }),
  },
  [IpcChannel.FileChooseSavePath]: {
    // O tipo viaja junto porque decide a extensão padrão: uma planilha gravada
    // como `.sdoc` abriria como documento vazio na próxima vez.
    request: z.object({
      suggestedName: z.string().min(1).max(255),
      kind: documentKindSchema.default('document'),
    }),
    response: saveResultSchema,
  },
  [IpcChannel.FileAutosave]: {
    // `path` nulo é trabalho que nunca foi gravado — o caso em que a
    // recuperação vale mais, porque não há arquivo nenhum a que voltar.
    request: z.object({
      path: z.string().nullable(),
      name: z.string().min(1).max(255),
      kind: documentKindSchema,
      content: z.string().max(MAX_TEXT_LENGTH),
    }),
    response: z.object({ savedAt: z.number().int() }),
  },
  [IpcChannel.RecoveryPeek]: {
    request: emptyRequest,
    response: z.object({ draft: draftSummarySchema.nullable() }),
  },
  [IpcChannel.RecoveryRestore]: {
    request: emptyRequest,
    response: z.object({
      draft: draftSummarySchema.extend({ content: z.string().max(MAX_TEXT_LENGTH) }).nullable(),
    }),
  },
  [IpcChannel.RecoveryDiscard]: {
    request: emptyRequest,
    response: z.object({ discarded: z.literal(true) }),
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
  [IpcChannel.PrintExportPdf]: {
    request: printRequestSchema.extend({ suggestedName: z.string().min(1).max(255) }),
    response: saveResultSchema,
  },
  [IpcChannel.PrintDialog]: {
    request: printRequestSchema,
    // `false` significa que o usuário cancelou — cancelar não é erro.
    response: z.object({ printed: z.boolean() }),
  },
  [IpcChannel.PrintPreview]: {
    request: printRequestSchema.extend({ title: z.string().max(255) }),
    response: z.object({ opened: z.literal(true) }),
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
