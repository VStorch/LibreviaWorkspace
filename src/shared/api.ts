import type { IpcChannel } from './ipc-channels.js'
import type { IpcRequest, IpcResponse, IpcResult } from './ipc.js'
import type { MenuCommand } from './types.js'

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
  [C in Exclude<IpcChannel, typeof IpcChannel.MenuCommand>]: IpcRequest<C>
}
type IpcResponseMap = {
  [C in Exclude<IpcChannel, typeof IpcChannel.MenuCommand>]: IpcResponse<C>
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
  }
  readonly recent: {
    list: Call<typeof IpcChannel.RecentList>
    clear: Call<typeof IpcChannel.RecentClear>
  }
  readonly image: {
    pick: Call<typeof IpcChannel.ImagePick>
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
}
