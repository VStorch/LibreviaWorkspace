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
  /**
   * Abre o diálogo "salvar como" e apenas **autoriza** o destino escolhido,
   * sem gravar. Separar escolha de gravação é o que permite avisar sobre perda
   * de formatação antes de qualquer byte tocar o disco.
   */
  FileChooseSavePath: 'file:choose-save-path',

  /**
   * Guarda o que está na tela como rascunho de recuperação.
   *
   * Não grava no arquivo do usuário: escrever por cima dele sozinho
   * transformaria "não salvei" em "salvei sem querer".
   */
  FileAutosave: 'file:autosave',

  /** Há rascunho de uma sessão que não terminou bem? Só os dados do aviso. */
  RecoveryPeek: 'recovery:peek',
  /** Devolve o conteúdo do rascunho e reata o vínculo com o arquivo original. */
  RecoveryRestore: 'recovery:restore',
  RecoveryDiscard: 'recovery:discard',

  RecentList: 'recent:list',
  RecentClear: 'recent:clear',

  /** Escolhe uma imagem no disco e devolve como data URI já validado. */
  ImagePick: 'image:pick',

  /** Gera o PDF e grava no destino escolhido pelo usuário. */
  PrintExportPdf: 'print:export-pdf',
  /** Abre o diálogo de impressão do sistema. */
  PrintDialog: 'print:dialog',
  /** Gera o PDF e abre numa janela de visualização. */
  PrintPreview: 'print:preview',

  /** Aviso nativo de alterações não salvas, reutilizado pelo renderer. */
  DialogConfirmDiscard: 'dialog:confirm-discard',
  /** Aviso de que salvar em .txt descarta a formatação. */
  DialogConfirmPlainText: 'dialog:confirm-plain-text',

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
  IpcChannel.FileChooseSavePath,
  IpcChannel.FileAutosave,
  IpcChannel.RecoveryPeek,
  IpcChannel.RecoveryRestore,
  IpcChannel.RecoveryDiscard,
  IpcChannel.RecentList,
  IpcChannel.RecentClear,
  IpcChannel.ImagePick,
  IpcChannel.PrintExportPdf,
  IpcChannel.PrintDialog,
  IpcChannel.PrintPreview,
  IpcChannel.DialogConfirmDiscard,
  IpcChannel.DialogConfirmPlainText,
  IpcChannel.WindowSetState,
  IpcChannel.WindowClose,
] as const

export type InvocableIpcChannel = (typeof INVOCABLE_IPC_CHANNELS)[number]
