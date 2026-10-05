import { z } from 'zod'
import {
  IpcChannel,
  INVOCABLE_IPC_CHANNELS,
  PUSH_IPC_CHANNELS,
  type InvocableIpcChannel,
  type PushIpcChannel,
} from './ipc-channels.js'
import {
  contextMenuTargetSchema,
  editorPreferencesPatchSchema,
  editorPreferencesSchema,
  pageSetupSchema,
} from './schemas.js'
import { DictionaryScope, EditCommand, MenuCommand, RevisionView } from './types.js'
import type { SerializedError } from './errors.js'
import {
  MAX_FILE_NAME_LENGTH,
  MAX_FONT_FAMILIES,
  MAX_FONT_FAMILY_LENGTH,
  MAX_NAME_LENGTH,
  MAX_PATH_LENGTH,
  MAX_USER_TEMPLATES,
} from './limits.js'

/**
 * The boundary is **language-neutral**: only serializable data crosses it, never Node objects,
 * shared Buffers, classes or third-party library types.
 */

/**
 * A `.sdoc` embeds images as data URIs, so it is much larger than the equivalent `.txt`. The cap
 * keeps an absurd file from freezing the UI; it does not restrict legitimate use.
 */
export const MAX_TEXT_LENGTH = 50_000_000

const documentKindSchema = z.enum(['document', 'spreadsheet'])

/** See `LossInventory` in types.ts: invisible and lost are not the same thing. */
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
  // The file is a Word template: the renderer opens it as a new, untitled document, and `path` is
  // only the package origin, never a save destination.
  template: z.boolean().optional(),
})

/** Builtin templates by id, the user's by path. */
const templateEntrySchema = z.object({
  source: z.enum(['builtin', 'user']),
  id: z.string().min(1).max(MAX_PATH_LENGTH),
  name: z.string().max(MAX_FILE_NAME_LENGTH),
  description: z.string().max(500),
})

const recentFileSchema = z.object({
  path: z.string(),
  name: z.string(),
  kind: documentKindSchema,
  openedAt: z.number().int(),
})

/** The recovery prompt shows which file it came from and when. */
const draftSummarySchema = z.object({
  path: z.string().nullable(),
  name: z.string(),
  kind: documentKindSchema,
  savedAt: z.number().int(),
})

const emptyRequest = z.object({})

const printRequestSchema = z.object({
  html: z.string().max(MAX_TEXT_LENGTH),
  page: pageSetupSchema,
  /**
   * The HTML is already split into paper-sized sheets.
   *
   * When true, `printToPDF` gets no margins and no header or footer: the page draws them itself. A
   * spreadsheet is not paginated: Chromium paginates the continuous table.
   */
  paged: z.boolean().default(false),
})

/** A canceled dialog is not an error: it is an expected outcome. */
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
      // The path the edited document was loaded from, `null` for a new document. Main only compares
      // it with the `.docx` it opened, to decide whether to write over that package or a new one:
      // nothing is read from it.
      origin: z.string().min(1).nullable(),
    }),
    response: z.object({
      path: z.string(),
      name: z.string(),
      inventory: inventorySchema.optional(),
    }),
  },
  [IpcChannel.TemplateList]: {
    request: emptyRequest,
    response: z.object({
      builtin: z.array(templateEntrySchema).max(100),
      user: z.array(templateEntrySchema).max(MAX_USER_TEMPLATES),
      folder: z.string(),
    }),
  },
  [IpcChannel.TemplateOpen]: {
    // Only source and id: main checks the id against the list it builds itself, so the renderer
    // cannot open an arbitrary path through here.
    request: z.object({ source: z.enum(['builtin', 'user']), id: z.string().min(1).max(MAX_PATH_LENGTH) }),
    response: z.object({ file: loadedFileSchema }),
  },
  [IpcChannel.TemplateBrowse]: {
    request: emptyRequest,
    response: openResultSchema,
  },
  [IpcChannel.TemplateOpenFolder]: {
    request: emptyRequest,
    response: z.object({ folder: z.string() }),
  },
  [IpcChannel.FileChooseSavePath]: {
    // The kind decides the default extension: a spreadsheet saved as `.sdoc` would open as an empty
    // document next time.
    request: z.object({
      suggestedName: z.string().min(1).max(MAX_FILE_NAME_LENGTH),
      kind: documentKindSchema.default('document'),
    }),
    response: saveResultSchema,
  },
  [IpcChannel.FileAutosave]: {
    // A null `path` is work that was never saved, where recovery matters most, since there is no
    // file to go back to.
    request: z.object({
      path: z.string().nullable(),
      name: z.string().min(1).max(MAX_FILE_NAME_LENGTH),
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
        // Already validated by byte signature in main.
        dataUrl: z.string(),
        name: z.string(),
      }),
    ]),
  },
  [IpcChannel.FontsList]: {
    request: emptyRequest,
    // Generous, but still a cap: a print shop machine passes a thousand families, and no font name
    // has a hundred characters. The limit protects the UI from broken system output.
    response: z.object({
      families: z.array(z.string().min(1).max(MAX_FONT_FAMILY_LENGTH)).max(MAX_FONT_FAMILIES),
    }),
  },
  [IpcChannel.PrintExportPdf]: {
    request: printRequestSchema.extend({ suggestedName: z.string().min(1).max(MAX_FILE_NAME_LENGTH) }),
    response: saveResultSchema,
  },
  [IpcChannel.FileExport]: {
    request: z.object({
      format: z.enum(['html', 'markdown', 'odt']),
      // Serialized as when saving: main builds the file from the model, not from ready HTML the
      // renderer could have swapped.
      content: z.string().max(MAX_TEXT_LENGTH),
      suggestedName: z.string().min(1).max(MAX_FILE_NAME_LENGTH),
    }),
    response: saveResultSchema,
  },
  [IpcChannel.PrintDialog]: {
    request: printRequestSchema,
    // `false` means the user canceled, which is not an error.
    response: z.object({ printed: z.boolean() }),
  },
  [IpcChannel.PrintPreview]: {
    request: printRequestSchema.extend({ title: z.string().max(MAX_FILE_NAME_LENGTH) }),
    response: z.object({ opened: z.literal(true) }),
  },
  [IpcChannel.DialogConfirmDiscard]: {
    request: z.object({ fileName: z.string().min(1).max(MAX_FILE_NAME_LENGTH) }),
    response: z.object({ choice: z.enum(['save', 'discard', 'cancel']) }),
  },
  [IpcChannel.DialogConfirmPlainText]: {
    request: z.object({ fileName: z.string().min(1).max(MAX_FILE_NAME_LENGTH) }),
    response: z.object({ choice: z.enum(['keep-plain', 'save-as-document', 'cancel']) }),
  },
  [IpcChannel.WindowSetState]: {
    request: z.object({
      title: z.string().max(300),
      isDirty: z.boolean(),
      /** For the check mark in the Review menu. */
      trackChanges: z.boolean(),
      /** For the checked item in Review → Show. */
      revisionView: z.enum(RevisionView),
    }),
    response: z.object({ applied: z.literal(true) }),
  },
  [IpcChannel.WindowClose]: {
    request: emptyRequest,
    response: z.object({ closing: z.literal(true) }),
  },
  [IpcChannel.WindowReady]: {
    request: emptyRequest,
    response: z.object({ applied: z.literal(true) }),
  },
  [IpcChannel.PreferencesGet]: {
    request: emptyRequest,
    response: editorPreferencesSchema,
  },
  [IpcChannel.PreferencesSet]: {
    // A patch, not the whole set: whoever toggles formatting marks has no opinion about spelling,
    // and sending all three back would let one click undo what another just turned on.
    request: editorPreferencesPatchSchema,
    // The resulting state, so the renderer does not have to guess it.
    response: editorPreferencesSchema,
  },
  [IpcChannel.EditCommandRun]: {
    request: z.object({ command: z.enum(EditCommand) }),
    response: z.object({ done: z.literal(true) }),
  },
  [IpcChannel.ClipboardReadText]: {
    request: emptyRequest,
    response: z.object({ text: z.string().max(MAX_TEXT_LENGTH) }),
  },
  [IpcChannel.SpellReplaceWord]: {
    request: z.object({ word: z.string().min(1).max(MAX_NAME_LENGTH) }),
    response: z.object({ replaced: z.literal(true) }),
  },
  [IpcChannel.SpellAddWord]: {
    request: z.object({ word: z.string().min(1).max(MAX_NAME_LENGTH), scope: z.enum(DictionaryScope) }),
    // `false` when the spellchecker refused the word: it is off, or the word has a character the
    // user dictionary does not accept. Not an error.
    response: z.object({ added: z.boolean() }),
  },
} as const

/**
 * The opposite direction: what main pushes to the renderer.
 *
 * The renderer validates with the **same** schema before acting. It looks excessive since main is
 * the sender, but it keeps both sides agreeing on the message shape, and one day one of them will
 * be rewritten without the other.
 */
export const pushContracts = {
  [IpcChannel.MenuCommand]: z.object({
    command: z.enum(MenuCommand),
    /** Only with "open recent". */
    path: z.string().optional(),
  }),
  [IpcChannel.ContextMenuRequested]: contextMenuTargetSchema,
  [IpcChannel.PreferencesChanged]: editorPreferencesSchema,
} as const

export type PushContracts = typeof pushContracts

export type PushPayload<C extends PushIpcChannel> = z.infer<PushContracts[C]>

export type IpcContracts = typeof ipcContracts

export type IpcRequest<C extends InvocableIpcChannel> = z.infer<IpcContracts[C]['request']>
export type IpcResponse<C extends InvocableIpcChannel> = z.infer<IpcContracts[C]['response']>

/**
 * Handlers never propagate exceptions through IPC: every call returns success or an already
 * sanitized error.
 */
export type IpcResult<T> =
  { readonly ok: true; readonly data: T } | { readonly ok: false; readonly error: SerializedError }

export { INVOCABLE_IPC_CHANNELS, PUSH_IPC_CHANNELS }
