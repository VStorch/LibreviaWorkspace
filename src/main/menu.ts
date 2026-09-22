import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { SHORTCUTS, acceleratorOf } from '@shared/shortcuts.js'
import { TABLE_ACTIONS, TableAction } from '@shared/table-actions.js'
import { MenuCommand } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
import { editorPreferences, updatePreferences } from './preferences.js'
import { devServerUrl, sendMenuCommand } from './window.js'

const isMac = process.platform === 'darwin'

function focusedWindow(): BrowserWindow | null {
  return BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0] ?? null
}

function dispatch(command: MenuCommand, path?: string): void {
  const window = focusedWindow()
  if (window === null) return
  sendMenuCommand(window, path === undefined ? { command } : { command, path })
}

async function buildRecentSubmenu(): Promise<MenuItemConstructorOptions[]> {
  const recent = await listRecentFiles()
  if (recent.length === 0) {
    return [{ label: 'Nenhum arquivo recente', enabled: false }]
  }

  return [
    ...recent.map<MenuItemConstructorOptions>((file) => ({
      label: file.name,
      // O caminho completo é útil quando há dois arquivos de mesmo nome em
      // pastas diferentes — situação comum em rede.
      toolTip: file.path,
      click: () => dispatch(MenuCommand.OpenRecent, file.path),
    })),
    { type: 'separator' },
    { label: 'Limpar recentes', click: () => dispatch(MenuCommand.ClearRecent) },
  ]
}

/**
 * O menu "Tabela", montado da lista única de ações.
 *
 * Os itens não se apagam fora de uma tabela: o menu nativo mora no processo main
 * e não sabe onde está o cursor, e reconstruí-lo a cada movimento dele custaria
 * mais do que vale. Fora de uma tabela os comandos do TableKit simplesmente não
 * fazem nada — e o menu de contexto, que sabe, só os oferece dentro de uma.
 */
function buildTableSubmenu(): MenuItemConstructorOptions[] {
  const items: MenuItemConstructorOptions[] = []
  let group = TABLE_ACTIONS[0]?.group

  for (const action of TABLE_ACTIONS) {
    if (action.group !== group) items.push({ type: 'separator' })
    group = action.group

    // `TableAction` é um subconjunto de `MenuCommand`, com os mesmos valores:
    // o `App` repassa ao editor pelo nome.
    const command: MenuCommand = action.id
    items.push({
      label: action.label,
      ...(action.id === TableAction.Insert ? { accelerator: acceleratorOf(SHORTCUTS.insertTable) } : {}),
      click: () => dispatch(command),
    })
  }

  return items
}

/**
 * Menu da Fase 1.
 *
 * Só entram itens que funcionam. Impressão (Fase 3), formatação (Fase 2) e
 * localizar/substituir (Fase 2) aparecerão junto com suas fases — um item de
 * menu desabilitado ou que não faz nada é pior que a ausência dele.
 */
async function buildTemplate(): Promise<MenuItemConstructorOptions[]> {
  const macAppMenu: MenuItemConstructorOptions[] = isMac
    ? [
        {
          label: APP_NAME,
          submenu: [
            { role: 'about', label: `Sobre o ${APP_NAME}` },
            { type: 'separator' },
            { role: 'hide', label: `Ocultar ${APP_NAME}` },
            { role: 'hideOthers', label: 'Ocultar outros' },
            { role: 'unhide', label: 'Mostrar todos' },
            { type: 'separator' },
            { role: 'quit', label: `Encerrar ${APP_NAME}` },
          ],
        },
      ]
    : []

  const preferences = editorPreferences()

  const viewSubmenu: MenuItemConstructorOptions[] = [
    {
      label: 'Marcas de formatação',
      type: 'checkbox',
      checked: preferences.invisibleCharacters,
      // A tecla e o porquê dela estão na tabela de atalhos.
      accelerator: acceleratorOf(SHORTCUTS.formattingMarks),
      click: () => {
        updatePreferences({ invisibleCharacters: !preferences.invisibleCharacters })
      },
    },
    { type: 'separator' },
    { role: 'resetZoom', label: 'Tamanho normal' },
    // Ampliar sai do `Ctrl+Shift+=` do sobrescrito: ver a tabela de atalhos.
    { role: 'zoomIn', label: 'Ampliar', accelerator: acceleratorOf(SHORTCUTS.zoomIn) },
    { role: 'zoomOut', label: 'Reduzir' },
    { type: 'separator' },
    { role: 'togglefullscreen', label: 'Tela cheia' },
  ]

  if (devServerUrl() !== null) {
    viewSubmenu.push(
      { type: 'separator' },
      { role: 'reload', label: 'Recarregar', accelerator: acceleratorOf(SHORTCUTS.reload) },
      { role: 'toggleDevTools', label: 'Ferramentas do desenvolvedor' },
    )
  }

  return [
    ...macAppMenu,
    {
      label: 'Arquivo',
      submenu: [
        {
          label: 'Novo documento',
          accelerator: acceleratorOf(SHORTCUTS.newDocument),
          click: () => dispatch(MenuCommand.NewDocument),
        },
        {
          label: 'Nova planilha',
          accelerator: acceleratorOf(SHORTCUTS.newSpreadsheet),
          click: () => dispatch(MenuCommand.NewSpreadsheet),
        },
        { type: 'separator' },
        {
          label: 'Abrir…',
          accelerator: acceleratorOf(SHORTCUTS.open),
          click: () => dispatch(MenuCommand.Open),
        },
        { label: 'Abrir recente', submenu: await buildRecentSubmenu() },
        { type: 'separator' },
        {
          label: 'Salvar',
          accelerator: acceleratorOf(SHORTCUTS.save),
          click: () => dispatch(MenuCommand.Save),
        },
        {
          label: 'Salvar como…',
          accelerator: acceleratorOf(SHORTCUTS.saveAs),
          click: () => dispatch(MenuCommand.SaveAs),
        },
        { type: 'separator' },
        { label: 'Configuração de página…', click: () => dispatch(MenuCommand.PageSetup) },
        { label: 'Visualizar impressão', click: () => dispatch(MenuCommand.PrintPreview) },
        { label: 'Exportar para PDF…', click: () => dispatch(MenuCommand.ExportPdf) },
        {
          label: 'Imprimir…',
          accelerator: acceleratorOf(SHORTCUTS.print),
          click: () => dispatch(MenuCommand.Print),
        },
        { type: 'separator' },
        {
          label: 'Fechar arquivo',
          accelerator: acceleratorOf(SHORTCUTS.closeFile),
          click: () => dispatch(MenuCommand.CloseFile),
        },
        { role: 'quit', label: 'Sair' },
      ],
    },
    {
      label: 'Editar',
      submenu: [
        { role: 'undo', label: 'Desfazer' },
        { role: 'redo', label: 'Refazer' },
        { type: 'separator' },
        { role: 'cut', label: 'Recortar' },
        { role: 'copy', label: 'Copiar' },
        { role: 'paste', label: 'Colar' },
        {
          label: 'Colar sem formatação',
          accelerator: acceleratorOf(SHORTCUTS.pasteWithoutFormat),
          click: () => dispatch(MenuCommand.PasteWithoutFormat),
        },
        { role: 'selectAll', label: 'Selecionar tudo' },
        { type: 'separator' },
        {
          label: 'Localizar e substituir…',
          accelerator: acceleratorOf(SHORTCUTS.findReplace),
          click: () => dispatch(MenuCommand.FindReplace),
        },
      ],
    },
    {
      label: 'Formatar',
      submenu: [
        {
          label: 'Parágrafo…',
          click: () => dispatch(MenuCommand.ParagraphSetup),
        },
        {
          label: 'Imagem…',
          click: () => dispatch(MenuCommand.ImageProperties),
        },
      ],
    },
    { label: 'Tabela', submenu: buildTableSubmenu() },
    {
      label: 'Inserir',
      submenu: [
        {
          label: 'Quebra de página',
          accelerator: acceleratorOf(SHORTCUTS.insertPageBreak),
          click: () => dispatch(MenuCommand.InsertPageBreak),
        },
        {
          label: 'Caractere especial…',
          click: () => dispatch(MenuCommand.SpecialCharacter),
        },
      ],
    },
    { label: 'Exibir', submenu: viewSubmenu },
    {
      label: 'Ferramentas',
      submenu: [
        {
          label: 'Verificação ortográfica',
          type: 'checkbox',
          checked: preferences.spellcheck,
          click: () => {
            updatePreferences({ spellcheck: !preferences.spellcheck })
          },
        },
        {
          label: 'Autocorreção tipográfica',
          type: 'checkbox',
          checked: preferences.typography,
          click: () => {
            updatePreferences({ typography: !preferences.typography })
          },
        },
        { type: 'separator' },
        {
          label: 'Contar palavras…',
          accelerator: acceleratorOf(SHORTCUTS.wordCount),
          click: () => dispatch(MenuCommand.WordCount),
        },
      ],
    },
    {
      label: 'Ajuda',
      submenu: [
        {
          label: `Sobre o ${APP_NAME}`,
          click: () => {
            const window = focusedWindow()
            if (window !== null) showAboutDialog(window, APP_NAME, app.getVersion())
          },
        },
      ],
    },
  ]
}

/** Reconstrói o menu. Necessário sempre que a lista de recentes mudar. */
export async function refreshMenu(): Promise<void> {
  Menu.setApplicationMenu(Menu.buildFromTemplate(await buildTemplate()))
}
