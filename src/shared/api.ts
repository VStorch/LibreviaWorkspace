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

/** Payload de um comando de menu. `path` só vem em "abrir recente". */
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
  /** Modelos do Word. */
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
    /** HTML ou Markdown. */
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
  readonly menu: {
    /** Assina os comandos do menu nativo. Devolve a função de cancelamento. */
    onCommand(listener: (payload: MenuCommandPayload) => void): () => void
  }
  readonly preferences: {
    get: Call<typeof IpcChannel.PreferencesGet>
    set: Call<typeof IpcChannel.PreferencesSet>
    /** Sem o aviso de volta, clicar no ¶ da barra deixaria o item do menu desmarcado. */
    onChange(listener: (preferences: EditorPreferences) => void): () => void
  }
  readonly edit: {
    /** Recortar, copiar e colar — os que dependem da área de transferência. */
    run: Call<typeof IpcChannel.EditCommandRun>
    /** O texto da área de transferência, para colar sem formatação. */
    readClipboardText: Call<typeof IpcChannel.ClipboardReadText>
  }
  readonly spell: {
    /** Troca a palavra errada debaixo do cursor pela escolhida. */
    replace: Call<typeof IpcChannel.SpellReplaceWord>
    addWord: Call<typeof IpcChannel.SpellAddWord>
  }
  readonly contextMenu: {
    /** O evento nasce no main, onde o corretor do Chromium diz a palavra errada e as sugestões. */
    onRequest(listener: (target: ContextMenuTarget) => void): () => void
  }
}
