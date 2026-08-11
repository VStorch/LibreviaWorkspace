/**
 * Nomes de canal IPC.
 *
 * Este arquivo é deliberadamente livre de dependências: ele é importado pelo
 * preload, que roda em contexto sandboxed e não pode carregar pacotes de
 * terceiros. Os schemas de validação ficam em `ipc.ts`, que só o main importa.
 */
export const IpcChannel = {
  /** Abre o diálogo do sistema e carrega o arquivo escolhido. */
  FileOpen: 'file:open',
  /** Carrega um caminho específico — só o que já estiver na lista de recentes. */
  FileOpenRecent: 'file:open-recent',
  /** Grava sobre um caminho já autorizado nesta sessão. */
  FileSave: 'file:save',
  /** Abre o diálogo "salvar como" e grava no destino escolhido. */
  FileSaveAs: 'file:save-as',

  RecentList: 'recent:list',
  RecentClear: 'recent:clear',

  /** Aviso nativo de alterações não salvas, reutilizado pelo renderer. */
  DialogConfirmDiscard: 'dialog:confirm-discard',

  /** Informa ao main o título e o estado de alterações não salvas. */
  WindowSetState: 'window:set-state',
  /** Pedido explícito de fechamento, já resolvido do lado do renderer. */
  WindowClose: 'window:close',

  /** Canal main → renderer: comandos disparados pelo menu nativo. */
  MenuCommand: 'menu:command',
} as const

export type IpcChannel = (typeof IpcChannel)[keyof typeof IpcChannel]

/** Canais no sentido renderer → main. O `MenuCommand` vai no sentido oposto. */
export const INVOCABLE_IPC_CHANNELS = [
  IpcChannel.FileOpen,
  IpcChannel.FileOpenRecent,
  IpcChannel.FileSave,
  IpcChannel.FileSaveAs,
  IpcChannel.RecentList,
  IpcChannel.RecentClear,
  IpcChannel.DialogConfirmDiscard,
  IpcChannel.WindowSetState,
  IpcChannel.WindowClose,
] as const

export type InvocableIpcChannel = (typeof INVOCABLE_IPC_CHANNELS)[number]
