import { contextBridge, ipcRenderer, type IpcRendererEvent } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { AppApi, MenuCommandPayload } from '@shared/api.js'
import type { ContextMenuTarget, EditorPreferences } from '@shared/types.js'

/**
 * Assinatura de um canal main → renderer.
 *
 * Escrita uma vez porque já são três: o `IpcRendererEvent` carrega referências ao
 * sistema de mensagens e não pode vazar para o renderer, e repetir esse cuidado
 * em cada assinante é como um dia ele deixaria de ser feito.
 *
 * Aqui não há validação de schema de propósito: o preload roda sandboxed e não
 * carrega zod. Quem valida a mensagem recebida é o renderer, com o mesmo
 * `pushContracts` que o main usou para mandá-la.
 */
function subscribe<T>(channel: string, listener: (payload: T) => void): () => void {
  const wrapped = (_event: IpcRendererEvent, payload: T): void => listener(payload)
  ipcRenderer.on(channel, wrapped)
  return () => {
    ipcRenderer.removeListener(channel, wrapped)
  }
}

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
    autosave: (payload) => ipcRenderer.invoke(IpcChannel.FileAutosave, payload),
  },
  recovery: {
    peek: (payload) => ipcRenderer.invoke(IpcChannel.RecoveryPeek, payload),
    restore: (payload) => ipcRenderer.invoke(IpcChannel.RecoveryRestore, payload),
    discard: (payload) => ipcRenderer.invoke(IpcChannel.RecoveryDiscard, payload),
  },
  recent: {
    list: (payload) => ipcRenderer.invoke(IpcChannel.RecentList, payload),
    clear: (payload) => ipcRenderer.invoke(IpcChannel.RecentClear, payload),
  },
  image: {
    pick: (payload) => ipcRenderer.invoke(IpcChannel.ImagePick, payload),
  },
  fonts: {
    list: (payload) => ipcRenderer.invoke(IpcChannel.FontsList, payload),
  },
  print: {
    exportPdf: (payload) => ipcRenderer.invoke(IpcChannel.PrintExportPdf, payload),
    dialog: (payload) => ipcRenderer.invoke(IpcChannel.PrintDialog, payload),
    preview: (payload) => ipcRenderer.invoke(IpcChannel.PrintPreview, payload),
  },
  dialog: {
    confirmDiscard: (payload) => ipcRenderer.invoke(IpcChannel.DialogConfirmDiscard, payload),
    confirmPlainText: (payload) => ipcRenderer.invoke(IpcChannel.DialogConfirmPlainText, payload),
  },
  window: {
    ready: (payload) => ipcRenderer.invoke(IpcChannel.WindowReady, payload),
    setState: (payload) => ipcRenderer.invoke(IpcChannel.WindowSetState, payload),
    close: (payload) => ipcRenderer.invoke(IpcChannel.WindowClose, payload),
  },
  menu: {
    onCommand: (listener) => subscribe<MenuCommandPayload>(IpcChannel.MenuCommand, listener),
  },
  preferences: {
    get: (payload) => ipcRenderer.invoke(IpcChannel.PreferencesGet, payload),
    set: (payload) => ipcRenderer.invoke(IpcChannel.PreferencesSet, payload),
    onChange: (listener) => subscribe<EditorPreferences>(IpcChannel.PreferencesChanged, listener),
  },
  edit: {
    run: (payload) => ipcRenderer.invoke(IpcChannel.EditCommandRun, payload),
    readClipboardText: (payload) => ipcRenderer.invoke(IpcChannel.ClipboardReadText, payload),
  },
  spell: {
    replace: (payload) => ipcRenderer.invoke(IpcChannel.SpellReplaceWord, payload),
    addWord: (payload) => ipcRenderer.invoke(IpcChannel.SpellAddWord, payload),
  },
  contextMenu: {
    onRequest: (listener) => subscribe<ContextMenuTarget>(IpcChannel.ContextMenuRequested, listener),
  },
}

contextBridge.exposeInMainWorld('api', api)
