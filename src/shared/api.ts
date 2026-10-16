import type { IpcChannel, PushIpcChannel } from './ipc-channels.js'
import type { IpcRequest, IpcResponse, IpcResult } from './ipc.js'
import type { ContextMenuTarget, EditorPreferences, MenuCommand } from './types.js'

type Call<C extends keyof IpcRequestMap> = (
  payload: IpcRequestMap[C],
) => Promise<IpcResult<IpcResponseMap[C]>>

type IpcRequestMap = {
  [C in Exclude<IpcChannel, PushIpcChannel>]: IpcRequest<C>
}
type IpcResponseMap = {
  [C in Exclude<IpcChannel, PushIpcChannel>]: IpcResponse<C>
}

export interface AiStreamEvent {
  readonly requestId: string
  readonly delta: string
}

/** `path` only comes with "open recent". */
export interface MenuCommandPayload {
  readonly command: MenuCommand
  readonly path?: string
}

export interface AppApi {
  readonly file: {
    open: Call<typeof IpcChannel.FileOpen>
    openRecent: Call<typeof IpcChannel.FileOpenRecent>
    save: Call<typeof IpcChannel.FileSave>
    chooseSavePath: Call<typeof IpcChannel.FileChooseSavePath>
    autosave: Call<typeof IpcChannel.FileAutosave>
  }
  readonly template: {
    list: Call<typeof IpcChannel.TemplateList>
    open: Call<typeof IpcChannel.TemplateOpen>
    browse: Call<typeof IpcChannel.TemplateBrowse>
    openFolder: Call<typeof IpcChannel.TemplateOpenFolder>
  }
  readonly recovery: {
    peek: Call<typeof IpcChannel.RecoveryPeek>
    restore: Call<typeof IpcChannel.RecoveryRestore>
    discard: Call<typeof IpcChannel.RecoveryDiscard>
  }
  readonly recent: {
    list: Call<typeof IpcChannel.RecentList>
    clear: Call<typeof IpcChannel.RecentClear>
  }
  readonly image: {
    pick: Call<typeof IpcChannel.ImagePick>
  }
  readonly fonts: {
    list: Call<typeof IpcChannel.FontsList>
  }
  readonly print: {
    exportPdf: Call<typeof IpcChannel.PrintExportPdf>
    exportDocument: Call<typeof IpcChannel.FileExport>
    dialog: Call<typeof IpcChannel.PrintDialog>
    preview: Call<typeof IpcChannel.PrintPreview>
  }
  readonly dialog: {
    confirmDiscard: Call<typeof IpcChannel.DialogConfirmDiscard>
    confirmPlainText: Call<typeof IpcChannel.DialogConfirmPlainText>
  }
  readonly window: {
    ready: Call<typeof IpcChannel.WindowReady>
    setState: Call<typeof IpcChannel.WindowSetState>
    close: Call<typeof IpcChannel.WindowClose>
  }
  readonly pdf: {
    share: Call<typeof IpcChannel.PdfShare>
    print: Call<typeof IpcChannel.PdfPrint>
    properties: Call<typeof IpcChannel.PdfProperties>
    saveProperties: Call<typeof IpcChannel.PdfSaveProperties>
  }
  readonly privacy: {
    inspect: Call<typeof IpcChannel.PrivacyInspect>
    remove: Call<typeof IpcChannel.PrivacyRemove>
  }
  readonly ai: {
    listProviders: Call<typeof IpcChannel.AiProvidersList>
    saveProvider: Call<typeof IpcChannel.AiProviderSave>
    removeProvider: Call<typeof IpcChannel.AiProviderRemove>
    chooseModel: Call<typeof IpcChannel.AiModelChoose>
    savePreferences: Call<typeof IpcChannel.AiPreferencesSave>
    chooseAgentExecutable: Call<typeof IpcChannel.AiAgentChooseExecutable>
    discover: Call<typeof IpcChannel.AiDiscover>
    listModels: Call<typeof IpcChannel.AiModelsList>
    generate: Call<typeof IpcChannel.AiGenerate>
    chat: Call<typeof IpcChannel.AiChat>
    cancel: Call<typeof IpcChannel.AiCancel>
    onStream(listener: (event: AiStreamEvent) => void): () => void
  }
  readonly menu: {
    onCommand(listener: (payload: MenuCommandPayload) => void): () => void
  }
  readonly preferences: {
    get: Call<typeof IpcChannel.PreferencesGet>
    set: Call<typeof IpcChannel.PreferencesSet>
    /** Without the echo, clicking the toolbar ¶ would leave the menu item unchecked. */
    onChange(listener: (preferences: EditorPreferences) => void): () => void
  }
  readonly edit: {
    run: Call<typeof IpcChannel.EditCommandRun>
    readClipboardText: Call<typeof IpcChannel.ClipboardReadText>
  }
  readonly spell: {
    replace: Call<typeof IpcChannel.SpellReplaceWord>
    addWord: Call<typeof IpcChannel.SpellAddWord>
  }
  readonly contextMenu: {
    /**
     * Raised in main, where Chromium's spellchecker knows the misspelled word and its suggestions.
     */
    onRequest(listener: (target: ContextMenuTarget) => void): () => void
  }
}
