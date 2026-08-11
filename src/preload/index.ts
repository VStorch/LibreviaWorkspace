import { contextBridge, ipcRenderer } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { AppApi } from '@shared/api.js'

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
  app: {
    ping: (payload) => ipcRenderer.invoke(IpcChannel.AppPing, payload),
  },
}

contextBridge.exposeInMainWorld('api', api)
