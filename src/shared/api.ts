import type { IpcChannel, PushIpcChannel } from './ipc-channels.js'
import type { IpcRequest, IpcResponse, IpcResult } from './ipc.js'
import type { ContextMenuTarget, EditorPreferences, MenuCommand } from './types.js'

/**
 * Superfície completa que o renderer enxerga do mundo externo.
 *
 * É deliberadamente uma interface explícita, e não um tipo inferido do preload:
 * ampliar o que o renderer pode fazer tem de ser uma edição consciente deste
 * arquivo, não um efeito colateral de mexer na implementação.
 */

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
    dialog: Call<typeof IpcChannel.PrintDialog>
    preview: Call<typeof IpcChannel.PrintPreview>
  }
  readonly dialog: {
    confirmDiscard: Call<typeof IpcChannel.DialogConfirmDiscard>
    confirmPlainText: Call<typeof IpcChannel.DialogConfirmPlainText>
  }
  readonly window: {
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
    /**
     * Assina a mudança de preferência, venha do menu nativo ou da barra.
     *
     * Existe porque as duas pontas podem ligar a mesma coisa: sem o aviso de
     * volta, clicar no ¶ da barra deixaria o item do menu desmarcado.
     */
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
    /**
     * Assina o clique com o botão direito.
     *
     * O evento nasce no `webContents`, no main: é lá que o corretor do Chromium
     * diz qual palavra está errada e o que sugere. O menu em si é desenhado pelo
     * renderer, como o da planilha.
     */
    onRequest(listener: (target: ContextMenuTarget) => void): () => void
  }
}
