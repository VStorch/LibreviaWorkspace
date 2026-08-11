import { contextBridge, ipcRenderer, type IpcRendererEvent } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { AppApi, MenuCommandPayload } from '@shared/api.js'

/**
 * Ponte entre renderer e main.
 *
 * Regras que valem para sempre neste arquivo:
 *  - nada de expor `ipcRenderer` cru, nem um `invoke` genérico: isso devolveria
 *    ao renderer o poder que o contextIsolation acabou de tirar;
 *  - um método por operação prevista no contrato, e nada além disso;
 *  - sem lógica de negócio — este é um encaminhador, não uma camada.
 */
const api: AppApi = {
  file: {
    open: (payload) => ipcRenderer.invoke(IpcChannel.FileOpen, payload),
    openRecent: (payload) => ipcRenderer.invoke(IpcChannel.FileOpenRecent, payload),
    save: (payload) => ipcRenderer.invoke(IpcChannel.FileSave, payload),
    chooseSavePath: (payload) => ipcRenderer.invoke(IpcChannel.FileChooseSavePath, payload),
  },
  recent: {
    list: (payload) => ipcRenderer.invoke(IpcChannel.RecentList, payload),
    clear: (payload) => ipcRenderer.invoke(IpcChannel.RecentClear, payload),
  },
  image: {
    pick: (payload) => ipcRenderer.invoke(IpcChannel.ImagePick, payload),
  },
  dialog: {
    confirmDiscard: (payload) => ipcRenderer.invoke(IpcChannel.DialogConfirmDiscard, payload),
    confirmPlainText: (payload) => ipcRenderer.invoke(IpcChannel.DialogConfirmPlainText, payload),
  },
  window: {
    setState: (payload) => ipcRenderer.invoke(IpcChannel.WindowSetState, payload),
    close: (payload) => ipcRenderer.invoke(IpcChannel.WindowClose, payload),
  },
  menu: {
    onCommand: (listener) => {
      // O `IpcRendererEvent` carrega referências ao sistema de mensagens e não
      // pode vazar para o renderer: só a payload atravessa.
      const wrapped = (_event: IpcRendererEvent, payload: MenuCommandPayload): void => listener(payload)
      ipcRenderer.on(IpcChannel.MenuCommand, wrapped)
      return () => {
        ipcRenderer.removeListener(IpcChannel.MenuCommand, wrapped)
      }
    },
  },
}

contextBridge.exposeInMainWorld('api', api)
