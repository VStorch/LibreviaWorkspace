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

  /**
   * As famílias de fonte instaladas na máquina.
   *
   * Vem do main porque descobri-las é executar programa do sistema, e o renderer
   * não executa nada. Lista vazia é resposta legítima: num sistema sem
   * `fontconfig` a barra segue com as fontes que o instalador leva.
   */
  FontsList: 'fonts:list',

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

  /** Lê as preferências de edição — ortografia, marcas, tipografia. */
  PreferencesGet: 'prefs:get',
  /** Liga ou desliga uma preferência. O main é quem guarda e quem aplica. */
  PreferencesSet: 'prefs:set',

  /**
   * Recortar, copiar e colar de verdade.
   *
   * O renderer não alcança a área de transferência do sistema — e não deve: quem
   * a lê e escreve é o `webContents`, no main.
   */
  EditCommandRun: 'edit:command',
  /** O texto da área de transferência, para colar sem formatação. */
  ClipboardReadText: 'clipboard:read-text',

  /** Troca a palavra errada pela sugestão escolhida no menu de contexto. */
  SpellReplaceWord: 'spell:replace',
  /** Guarda a palavra no dicionário — para sempre ou só nesta sessão. */
  SpellAddWord: 'spell:add-word',

  /** Canal main → renderer: comandos disparados pelo menu nativo. */
  MenuCommand: 'menu:command',
  /** Canal main → renderer: o botão direito foi clicado, e sobre o quê. */
  ContextMenuRequested: 'context-menu:requested',
  /** Canal main → renderer: uma preferência mudou, venha de onde vier. */
  PreferencesChanged: 'prefs:changed',
} as const

export type IpcChannel = (typeof IpcChannel)[keyof typeof IpcChannel]

/** Canais no sentido renderer → main. Ver `PUSH_IPC_CHANNELS` para o oposto. */
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
  IpcChannel.PrintDialog,
  IpcChannel.PrintPreview,
  IpcChannel.DialogConfirmDiscard,
  IpcChannel.DialogConfirmPlainText,
  IpcChannel.WindowSetState,
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
 * Canais no sentido main → renderer.
 *
 * Ficam listados à parte porque não têm handler: o main empurra, o renderer
 * escuta. A lista existe para que o tipo da API do renderer saiba distinguir os
 * dois sentidos — antes havia um só, e o `Exclude` era escrito à mão em `api.ts`.
 */
export const PUSH_IPC_CHANNELS = [
  IpcChannel.MenuCommand,
  IpcChannel.ContextMenuRequested,
  IpcChannel.PreferencesChanged,
] as const

export type PushIpcChannel = (typeof PUSH_IPC_CHANNELS)[number]
