import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { APP_NAME } from '@shared/constants.js'
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
      /**
       * `Ctrl+F10`, e não o `Ctrl+*` do Word.
       *
       * `Ctrl+Shift+8` **é** o `Ctrl+*`, e é também o atalho da lista com
       * marcadores no Tiptap. Acelerador de menu é registrado no main e
       * intercepta a tecla antes de o renderer vê-la — pela mesma razão do zoom
       * logo abaixo, quem se muda é o item novo. `Ctrl+F10` é o que o LibreOffice
       * usa para isto.
       */
      accelerator: 'CmdOrCtrl+F10',
      click: () => {
        updatePreferences({ invisibleCharacters: !preferences.invisibleCharacters })
      },
    },
    { type: 'separator' },
    { role: 'resetZoom', label: 'Tamanho normal' },
    /**
     * Ampliar sai do `Ctrl+Shift+=`, e não por capricho.
     *
     * O acelerador padrão do papel `zoomIn` é `CommandOrControl+Plus`, e no
     * Electron "Plus" é a tecla do `=` **com Shift** — a mesma combinação que no
     * Word liga o sobrescrito. Acelerador de menu é registrado no processo main e
     * intercepta a tecla antes de o renderer vê-la: deixá-lo aqui faria o atalho
     * de sobrescrito nunca rodar, e atalho morto é pior do que atalho ausente.
     *
     * Num editor de texto a formatação vem antes do zoom, então quem se muda é o
     * zoom — para o `+` do teclado numérico, que não disputa com tecla nenhuma.
     * Reduzir fica onde estava: `Ctrl+-` não colide com nada.
     */
    { role: 'zoomIn', label: 'Ampliar', accelerator: 'CmdOrCtrl+numadd' },
    { role: 'zoomOut', label: 'Reduzir' },
    { type: 'separator' },
    { role: 'togglefullscreen', label: 'Tela cheia' },
  ]

  if (devServerUrl() !== null) {
    viewSubmenu.push(
      { type: 'separator' },
      // `Ctrl+Shift+R` e não `Ctrl+R`: em desenvolvimento o padrão do papel
      // `reload` engoliria o `Ctrl+R` de "alinhar à direita", e o atalho pareceria
      // quebrado só na máquina de quem programa.
      { role: 'reload', label: 'Recarregar', accelerator: 'CmdOrCtrl+Shift+R' },
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
          accelerator: 'CmdOrCtrl+N',
          click: () => dispatch(MenuCommand.NewDocument),
        },
        {
          label: 'Nova planilha',
          accelerator: 'CmdOrCtrl+Shift+N',
          click: () => dispatch(MenuCommand.NewSpreadsheet),
        },
        { type: 'separator' },
        { label: 'Abrir…', accelerator: 'CmdOrCtrl+O', click: () => dispatch(MenuCommand.Open) },
        { label: 'Abrir recente', submenu: await buildRecentSubmenu() },
        { type: 'separator' },
        { label: 'Salvar', accelerator: 'CmdOrCtrl+S', click: () => dispatch(MenuCommand.Save) },
        {
          label: 'Salvar como…',
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => dispatch(MenuCommand.SaveAs),
        },
        { type: 'separator' },
        { label: 'Configuração de página…', click: () => dispatch(MenuCommand.PageSetup) },
        { label: 'Visualizar impressão', click: () => dispatch(MenuCommand.PrintPreview) },
        { label: 'Exportar para PDF…', click: () => dispatch(MenuCommand.ExportPdf) },
        { label: 'Imprimir…', accelerator: 'CmdOrCtrl+P', click: () => dispatch(MenuCommand.Print) },
        { type: 'separator' },
        {
          label: 'Fechar arquivo',
          accelerator: 'CmdOrCtrl+W',
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
          // O Chromium já responde a `Ctrl+Shift+V` dentro de um campo editável,
          // e o que ele faz não é o que o Word faz. O acelerador daqui é
          // registrado no main e chega primeiro, então passa a valer o nosso.
          accelerator: 'CmdOrCtrl+Shift+V',
          click: () => dispatch(MenuCommand.PasteWithoutFormat),
        },
        { role: 'selectAll', label: 'Selecionar tudo' },
        { type: 'separator' },
        {
          label: 'Localizar e substituir…',
          accelerator: 'CmdOrCtrl+F',
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
      ],
    },
    {
      label: 'Inserir',
      submenu: [
        {
          label: 'Quebra de página',
          accelerator: 'CmdOrCtrl+Enter',
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
          // O mesmo atalho do Word.
          label: 'Contar palavras…',
          accelerator: 'CmdOrCtrl+Shift+G',
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
