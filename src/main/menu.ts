import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { MenuCommand } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
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

  const viewSubmenu: MenuItemConstructorOptions[] = [
    { role: 'resetZoom', label: 'Tamanho normal' },
    { role: 'zoomIn', label: 'Ampliar' },
    { role: 'zoomOut', label: 'Reduzir' },
    { type: 'separator' },
    { role: 'togglefullscreen', label: 'Tela cheia' },
  ]

  if (devServerUrl() !== null) {
    viewSubmenu.push(
      { type: 'separator' },
      { role: 'reload', label: 'Recarregar' },
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
          // Habilitado na Fase 5, quando existir editor de planilhas.
          enabled: false,
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
        { role: 'selectAll', label: 'Selecionar tudo' },
      ],
    },
    { label: 'Exibir', submenu: viewSubmenu },
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
