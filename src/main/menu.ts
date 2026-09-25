import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { LANGUAGES, LANGUAGE_NAMES, type MessageKey } from '@shared/i18n/index.js'
import { SHORTCUTS, acceleratorOf } from '@shared/shortcuts.js'
import { TABLE_ACTIONS, TableAction } from '@shared/table-actions.js'
import { MenuCommand, Theme } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
import { t } from './i18n.js'
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
    return [{ label: t('menu.file.noRecent'), enabled: false }]
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
    { label: t('menu.file.clearRecent'), click: () => dispatch(MenuCommand.ClearRecent) },
  ]
}

/** Um dos três temas, como item de rádio marcado conforme o que está escolhido. */
function themeItem(theme: Theme, key: MessageKey, chosen: Theme): MenuItemConstructorOptions {
  return {
    label: t(key),
    type: 'radio',
    checked: chosen === theme,
    click: () => {
      updatePreferences({ theme })
    },
  }
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
      label: t(action.labelKey),
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
            { role: 'about', label: t('menu.help.about') },
            { type: 'separator' },
            { role: 'hide', label: t('menu.app.hide') },
            { role: 'hideOthers', label: t('menu.app.hideOthers') },
            { role: 'unhide', label: t('menu.app.unhide') },
            { type: 'separator' },
            { role: 'quit', label: t('menu.app.quit') },
          ],
        },
      ]
    : []

  const preferences = editorPreferences()

  const viewSubmenu: MenuItemConstructorOptions[] = [
    {
      label: t('view.showToolbar'),
      type: 'checkbox',
      checked: preferences.showToolbar,
      click: () => {
        updatePreferences({ showToolbar: !preferences.showToolbar })
      },
    },
    {
      label: t('view.showStatusBar'),
      type: 'checkbox',
      checked: preferences.showStatusBar,
      click: () => {
        updatePreferences({ showStatusBar: !preferences.showStatusBar })
      },
    },
    { type: 'separator' },
    {
      label: t('view.reading'),
      type: 'checkbox',
      checked: preferences.readingMode,
      accelerator: acceleratorOf(SHORTCUTS.readingMode),
      click: () => {
        updatePreferences({ readingMode: !preferences.readingMode })
      },
    },
    { type: 'separator' },
    {
      label: t('menu.view.formattingMarks'),
      type: 'checkbox',
      checked: preferences.invisibleCharacters,
      // A tecla e o porquê dela estão na tabela de atalhos.
      accelerator: acceleratorOf(SHORTCUTS.formattingMarks),
      click: () => {
        updatePreferences({ invisibleCharacters: !preferences.invisibleCharacters })
      },
    },
    { type: 'separator' },
    {
      label: t('view.theme'),
      submenu: [
        // Botões de rádio, e não caixas: os três valores são exclusivos, e uma
        // caixa marcada em dois deles ao mesmo tempo não quer dizer nada.
        themeItem(Theme.System, 'view.theme.system', preferences.theme),
        themeItem(Theme.Light, 'view.theme.light', preferences.theme),
        themeItem(Theme.Dark, 'view.theme.dark', preferences.theme),
      ],
    },
    {
      label: t('view.language'),
      // Cada idioma escrito nele mesmo, e por isso sem passar pelo catálogo:
      // quem procura "English" num menu em português não acharia "Inglês".
      submenu: LANGUAGES.map<MenuItemConstructorOptions>((language) => ({
        label: LANGUAGE_NAMES[language],
        type: 'radio',
        checked: preferences.language === language,
        click: () => {
          updatePreferences({ language })
        },
      })),
    },
    { type: 'separator' },
    // O zoom é da folha, e não da janela: o do Chromium aumentava também as
    // barras e os diálogos. Quem calcula o degrau é o renderer, que sabe quanto
    // vale o "ajustar à largura" na janela de agora.
    {
      label: t('menu.view.resetZoom'),
      accelerator: acceleratorOf(SHORTCUTS.zoomReset),
      click: () => dispatch(MenuCommand.ZoomReset),
    },
    // Ampliar sai do `Ctrl+Shift+=` do sobrescrito: ver a tabela de atalhos.
    {
      label: t('menu.view.zoomIn'),
      accelerator: acceleratorOf(SHORTCUTS.zoomIn),
      click: () => dispatch(MenuCommand.ZoomIn),
    },
    {
      label: t('menu.view.zoomOut'),
      accelerator: acceleratorOf(SHORTCUTS.zoomOut),
      click: () => dispatch(MenuCommand.ZoomOut),
    },
    {
      label: t('menu.view.zoomFitWidth'),
      type: 'checkbox',
      checked: preferences.zoomFit,
      click: () => dispatch(MenuCommand.ZoomFitWidth),
    },
    { type: 'separator' },
    {
      role: 'togglefullscreen',
      label: t('menu.view.fullScreen'),
      accelerator: isMac ? 'Ctrl+Command+F' : 'F11',
    },
  ]

  if (devServerUrl() !== null) {
    viewSubmenu.push(
      { type: 'separator' },
      { role: 'reload', label: t('menu.view.reload'), accelerator: acceleratorOf(SHORTCUTS.reload) },
      { role: 'toggleDevTools', label: t('menu.view.devTools') },
    )
  }

  return [
    ...macAppMenu,
    {
      label: t('menu.file'),
      submenu: [
        {
          label: t('menu.file.newDocument'),
          accelerator: acceleratorOf(SHORTCUTS.newDocument),
          click: () => dispatch(MenuCommand.NewDocument),
        },
        {
          label: t('menu.file.newSpreadsheet'),
          accelerator: acceleratorOf(SHORTCUTS.newSpreadsheet),
          click: () => dispatch(MenuCommand.NewSpreadsheet),
        },
        { type: 'separator' },
        {
          label: t('menu.file.open'),
          accelerator: acceleratorOf(SHORTCUTS.open),
          click: () => dispatch(MenuCommand.Open),
        },
        { label: t('menu.file.openRecent'), submenu: await buildRecentSubmenu() },
        { type: 'separator' },
        {
          label: t('menu.file.save'),
          accelerator: acceleratorOf(SHORTCUTS.save),
          click: () => dispatch(MenuCommand.Save),
        },
        {
          label: t('menu.file.saveAs'),
          accelerator: acceleratorOf(SHORTCUTS.saveAs),
          click: () => dispatch(MenuCommand.SaveAs),
        },
        { type: 'separator' },
        { label: t('menu.file.pageSetup'), click: () => dispatch(MenuCommand.PageSetup) },
        { label: t('menu.file.printPreview'), click: () => dispatch(MenuCommand.PrintPreview) },
        { label: t('menu.file.exportPdf'), click: () => dispatch(MenuCommand.ExportPdf) },
        {
          label: t('menu.file.print'),
          accelerator: acceleratorOf(SHORTCUTS.print),
          click: () => dispatch(MenuCommand.Print),
        },
        { type: 'separator' },
        {
          label: t('menu.file.close'),
          accelerator: acceleratorOf(SHORTCUTS.closeFile),
          click: () => dispatch(MenuCommand.CloseFile),
        },
        { role: 'quit', label: t('menu.file.quit') },
      ],
    },
    {
      label: t('menu.edit'),
      submenu: [
        { role: 'undo', label: t('menu.edit.undo') },
        { role: 'redo', label: t('menu.edit.redo') },
        { type: 'separator' },
        { role: 'cut', label: t('menu.edit.cut') },
        { role: 'copy', label: t('menu.edit.copy') },
        { role: 'paste', label: t('menu.edit.paste') },
        {
          label: t('menu.edit.pasteWithoutFormat'),
          accelerator: acceleratorOf(SHORTCUTS.pasteWithoutFormat),
          click: () => dispatch(MenuCommand.PasteWithoutFormat),
        },
        { role: 'selectAll', label: t('menu.edit.selectAll') },
        { type: 'separator' },
        {
          label: t('menu.edit.findReplace'),
          accelerator: acceleratorOf(SHORTCUTS.findReplace),
          click: () => dispatch(MenuCommand.FindReplace),
        },
      ],
    },
    {
      label: t('menu.format'),
      submenu: [
        {
          label: t('menu.format.paragraph'),
          click: () => dispatch(MenuCommand.ParagraphSetup),
        },
        {
          label: t('menu.format.image'),
          click: () => dispatch(MenuCommand.ImageProperties),
        },
      ],
    },
    { label: t('menu.table'), submenu: buildTableSubmenu() },
    {
      label: t('menu.insert'),
      submenu: [
        {
          label: t('menu.insert.pageBreak'),
          accelerator: acceleratorOf(SHORTCUTS.insertPageBreak),
          click: () => dispatch(MenuCommand.InsertPageBreak),
        },
        {
          label: t('menu.insert.specialCharacter'),
          click: () => dispatch(MenuCommand.SpecialCharacter),
        },
      ],
    },
    { label: t('menu.view'), submenu: viewSubmenu },
    {
      label: t('menu.tools'),
      submenu: [
        {
          label: t('menu.tools.spellcheck'),
          type: 'checkbox',
          checked: preferences.spellcheck,
          click: () => {
            updatePreferences({ spellcheck: !preferences.spellcheck })
          },
        },
        {
          label: t('menu.tools.typography'),
          type: 'checkbox',
          checked: preferences.typography,
          click: () => {
            updatePreferences({ typography: !preferences.typography })
          },
        },
        { type: 'separator' },
        {
          label: t('menu.tools.wordCount'),
          accelerator: acceleratorOf(SHORTCUTS.wordCount),
          click: () => dispatch(MenuCommand.WordCount),
        },
      ],
    },
    {
      label: t('menu.help'),
      submenu: [
        {
          label: t('menu.help.about'),
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
