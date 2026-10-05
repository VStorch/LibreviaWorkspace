/**
 * Free of dependencies on purpose: the sandboxed preload imports it and cannot load third-party
 * packages. The validation schemas live in `ipc.ts`, which only main imports.
 */
export const IpcChannel = {
  FileOpen: 'file:open',
  /** A recent file, or one handed to main by the file manager. */
  FileOpenRecent: 'file:open-recent',
  /** Writes over a path already authorized in this session. */
  FileSave: 'file:save',
  /**
   * Only **authorizes** the chosen destination, without writing. Separating the choice from the
   * write is what allows warning about formatting loss before any byte reaches the disk.
   */
  FileChooseSavePath: 'file:choose-save-path',

  /**
   * Stores what is on screen as a recovery draft.
   *
   * Never writes to the user's file: overwriting it unprompted would turn "I didn't save" into "I
   * saved by accident".
   */
  FileAutosave: 'file:autosave',

  /** Only the data the prompt needs. */
  RecoveryPeek: 'recovery:peek',
  /** Also restores the link to the original file. */
  RecoveryRestore: 'recovery:restore',
  RecoveryDiscard: 'recovery:discard',

  RecentList: 'recent:list',
  RecentClear: 'recent:clear',

  /** Returns an already validated data URI. */
  ImagePick: 'image:pick',

  /**
   * Comes from main because discovering fonts means running a system program, and the renderer runs
   * nothing. An empty list is a valid answer: without `fontconfig` the toolbar keeps the bundled
   * fonts.
   */
  FontsList: 'fonts:list',

  PrintExportPdf: 'print:export-pdf',
  /** The document being edited keeps its path and its state. */
  FileExport: 'file:export',
  TemplateList: 'template:list',
  TemplateOpen: 'template:open',
  TemplateBrowse: 'template:browse',
  TemplateOpenFolder: 'template:open-folder',
  PrintDialog: 'print:dialog',
  PrintPreview: 'print:preview',

  DialogConfirmDiscard: 'dialog:confirm-discard',
  DialogConfirmPlainText: 'dialog:confirm-plain-text',

  WindowSetState: 'window:set-state',
  /** The renderer has subscribed to the commands and can receive files from the file manager. */
  WindowReady: 'window:ready',
  /** Already resolved on the renderer side. */
  WindowClose: 'window:close',

  PreferencesGet: 'prefs:get',
  /** Main stores and applies it. */
  PreferencesSet: 'prefs:set',

  /**
   * The renderer cannot reach the system clipboard, and should not: `webContents`, in main, reads
   * and writes it.
   */
  EditCommandRun: 'edit:command',
  ClipboardReadText: 'clipboard:read-text',

  SpellReplaceWord: 'spell:replace',
  /** Forever, or only for this session. */
  SpellAddWord: 'spell:add-word',

  MenuCommand: 'menu:command',
  ContextMenuRequested: 'context-menu:requested',
  /** Whatever its origin. */
  PreferencesChanged: 'prefs:changed',
} as const

export type IpcChannel = (typeof IpcChannel)[keyof typeof IpcChannel]

/** See `PUSH_IPC_CHANNELS` for the other direction. */
export const INVOCABLE_IPC_CHANNELS = [
  IpcChannel.FileOpen,
  IpcChannel.FileOpenRecent,
  IpcChannel.FileSave,
  IpcChannel.FileChooseSavePath,
  IpcChannel.FileAutosave,
  IpcChannel.RecoveryPeek,
  IpcChannel.RecoveryRestore,
  IpcChannel.RecoveryDiscard,
  IpcChannel.RecentList,
  IpcChannel.RecentClear,
  IpcChannel.ImagePick,
  IpcChannel.FontsList,
  IpcChannel.PrintExportPdf,
  IpcChannel.FileExport,
  IpcChannel.TemplateList,
  IpcChannel.TemplateOpen,
  IpcChannel.TemplateBrowse,
  IpcChannel.TemplateOpenFolder,
  IpcChannel.PrintDialog,
  IpcChannel.PrintPreview,
  IpcChannel.DialogConfirmDiscard,
  IpcChannel.DialogConfirmPlainText,
  IpcChannel.WindowSetState,
  IpcChannel.WindowReady,
  IpcChannel.WindowClose,
  IpcChannel.PreferencesGet,
  IpcChannel.PreferencesSet,
  IpcChannel.EditCommandRun,
  IpcChannel.ClipboardReadText,
  IpcChannel.SpellReplaceWord,
  IpcChannel.SpellAddWord,
] as const

export type InvocableIpcChannel = (typeof INVOCABLE_IPC_CHANNELS)[number]

/**
 * Listed apart because they have no handler: main pushes, the renderer listens. The list lets the
 * renderer API type tell the two directions apart.
 */
export const PUSH_IPC_CHANNELS = [
  IpcChannel.MenuCommand,
  IpcChannel.ContextMenuRequested,
  IpcChannel.PreferencesChanged,
] as const

export type PushIpcChannel = (typeof PUSH_IPC_CHANNELS)[number]
