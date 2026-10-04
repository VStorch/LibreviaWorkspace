import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { LANGUAGES, LANGUAGE_NAMES, type MessageKey } from '@shared/i18n/index.js'
import { SHORTCUTS, acceleratorOf } from '@shared/shortcuts.js'
import { TABLE_ACTIONS, TableAction } from '@shared/table-actions.js'
import { MenuCommand, RevisionView, Theme, type EditorPreferences } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
import { t } from './i18n.js'
import { editorPreferences, updatePreferences } from './preferences.js'
import { devServerUrl, sendMenuCommand } from './window.js'

const isMac = process.platform === 'darwin'

let trackChangesOn = false

export function setTrackChangesChecked(on: boolean): void {
  if (on === trackChangesOn) return
  trackChangesOn = on
  void refreshMenu()
}

let revisionView: RevisionView = RevisionView.All

export function setRevisionViewChecked(view: RevisionView): void {
  if (view === revisionView) return
  revisionView = view
  void refreshMenu()
}

function revisionViewItem(
  view: RevisionView,
  key: MessageKey,
  command: MenuCommand,
): MenuItemConstructorOptions {
  return {
    label: t(key),
    type: 'radio',
    checked: revisionView === view,
    click: () => dispatch(command),
  }
}

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
      // Dois arquivos de mesmo nome em pastas diferentes são comuns em rede.
      toolTip: file.path,
      click: () => dispatch(MenuCommand.OpenRecent, file.path),
    })),
    { type: 'separator' },
    commandItem('menu.file.clearRecent', MenuCommand.ClearRecent),
  ]
}

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

function commandItem(
  key: MessageKey,
  command: MenuCommand,
  accelerator?: string,
): MenuItemConstructorOptions {
  return {
    label: t(key),
    ...(accelerator === undefined ? {} : { accelerator }),
    click: () => dispatch(command),
  }
}

type BooleanPreference = {
  [K in keyof EditorPreferences]: EditorPreferences[K] extends boolean ? K : never
}[keyof EditorPreferences]

function preferenceToggle(
  key: MessageKey,
  preference: BooleanPreference,
  preferences: EditorPreferences,
  accelerator?: string,
): MenuItemConstructorOptions {
  return {
    label: t(key),
    type: 'checkbox',
    checked: preferences[preference],
    ...(accelerator === undefined ? {} : { accelerator }),
    click: () => {
      updatePreferences({ [preference]: !preferences[preference] })
    },
  }
}

/** O zoom é da folha; o degrau é do renderer, que sabe quanto vale o "ajustar à largura". */
function zoomItems(preferences: EditorPreferences): MenuItemConstructorOptions[] {
  return [
    commandItem('menu.view.resetZoom', MenuCommand.ZoomReset, acceleratorOf(SHORTCUTS.zoomReset)),
    commandItem('menu.view.zoomIn', MenuCommand.ZoomIn, acceleratorOf(SHORTCUTS.zoomIn)),
    commandItem('menu.view.zoomOut', MenuCommand.ZoomOut, acceleratorOf(SHORTCUTS.zoomOut)),
    {
      label: t('menu.view.zoomFitWidth'),
      type: 'checkbox',
      checked: preferences.zoomFit,
      click: () => dispatch(MenuCommand.ZoomFitWidth),
    },
  ]
}

/** Os itens não se apagam fora de uma tabela: o main não sabe onde está o cursor. */
function buildTableSubmenu(): MenuItemConstructorOptions[] {
  const items: MenuItemConstructorOptions[] = []
  let group = TABLE_ACTIONS[0]?.group

  for (const action of TABLE_ACTIONS) {
    if (action.group !== group) items.push({ type: 'separator' })
    group = action.group

    // `TableAction` é um subconjunto de `MenuCommand`, com os mesmos valores.
    const command: MenuCommand = action.id
    items.push({
      label: t(action.labelKey),
      ...(action.id === TableAction.Insert ? { accelerator: acceleratorOf(SHORTCUTS.insertTable) } : {}),
      click: () => dispatch(command),
    })
  }

  return items
}

/** Só entram itens que funcionam. */
async function buildTemplate(): Promise<MenuItemConstructorOptions[]> {
  const preferences = editorPreferences()

  return [
    ...macAppMenu(),
    fileMenu(await buildRecentSubmenu()),
    editMenu(),
    formatMenu(),
    { label: t('menu.table'), submenu: buildTableSubmenu() },
    insertMenu(),
    referencesMenu(),
    reviewMenu(),
    { label: t('menu.view'), submenu: viewSubmenu(preferences) },
    toolsMenu(preferences),
    helpMenu(),
  ]
}

function macAppMenu(): MenuItemConstructorOptions[] {
  return isMac
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
}

function viewSubmenu(preferences: EditorPreferences): MenuItemConstructorOptions[] {
  const items: MenuItemConstructorOptions[] = [
    preferenceToggle('view.showToolbar', 'showToolbar', preferences),
    preferenceToggle('view.showStatusBar', 'showStatusBar', preferences),
    { type: 'separator' },
    preferenceToggle('view.reading', 'readingMode', preferences, acceleratorOf(SHORTCUTS.readingMode)),
    { type: 'separator' },
    preferenceToggle(
      'menu.view.formattingMarks',
      'invisibleCharacters',
      preferences,
      acceleratorOf(SHORTCUTS.formattingMarks),
    ),
    preferenceToggle(
      'menu.view.navigationPane',
      'navigationPane',
      preferences,
      acceleratorOf(SHORTCUTS.navigationPane),
    ),
    preferenceToggle('menu.view.commentsPane', 'commentsPane', preferences),
    { type: 'separator' },
    {
      label: t('view.theme'),
      submenu: [
        themeItem(Theme.System, 'view.theme.system', preferences.theme),
        themeItem(Theme.Light, 'view.theme.light', preferences.theme),
        themeItem(Theme.Dark, 'view.theme.dark', preferences.theme),
      ],
    },
    {
      label: t('view.language'),
      // Cada idioma escrito nele mesmo, fora do catálogo: quem procura "English" não acharia "Inglês".
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
    ...zoomItems(preferences),
    { type: 'separator' },
    {
      role: 'togglefullscreen',
      label: t('menu.view.fullScreen'),
      accelerator: isMac ? 'Ctrl+Command+F' : 'F11',
    },
  ]

  if (devServerUrl() !== null) {
    items.push(
      { type: 'separator' },
      { role: 'reload', label: t('menu.view.reload'), accelerator: acceleratorOf(SHORTCUTS.reload) },
      { role: 'toggleDevTools', label: t('menu.view.devTools') },
    )
  }
  return items
}

function fileMenu(recent: MenuItemConstructorOptions[]): MenuItemConstructorOptions {
  return {
    label: t('menu.file'),
    submenu: [
      commandItem('menu.file.newDocument', MenuCommand.NewDocument, acceleratorOf(SHORTCUTS.newDocument)),
      commandItem(
        'menu.file.newSpreadsheet',
        MenuCommand.NewSpreadsheet,
        acceleratorOf(SHORTCUTS.newSpreadsheet),
      ),
      commandItem('menu.file.newFromTemplate', MenuCommand.NewFromTemplate),
      { type: 'separator' },
      commandItem('menu.file.open', MenuCommand.Open, acceleratorOf(SHORTCUTS.open)),
      { label: t('menu.file.openRecent'), submenu: recent },
      { type: 'separator' },
      commandItem('menu.file.save', MenuCommand.Save, acceleratorOf(SHORTCUTS.save)),
      commandItem('menu.file.saveAs', MenuCommand.SaveAs, acceleratorOf(SHORTCUTS.saveAs)),
      { type: 'separator' },
      commandItem('menu.file.pageSetup', MenuCommand.PageSetup),
      commandItem('menu.file.printPreview', MenuCommand.PrintPreview),
      commandItem('menu.file.exportPdf', MenuCommand.ExportPdf),
      {
        label: t('menu.file.exportAs'),
        submenu: [
          commandItem('menu.file.exportHtml', MenuCommand.ExportHtml),
          commandItem('menu.file.exportMarkdown', MenuCommand.ExportMarkdown),
          commandItem('menu.file.exportOdt', MenuCommand.ExportOdt),
        ],
      },
      commandItem('menu.file.properties', MenuCommand.DocumentProperties),
      commandItem('menu.file.print', MenuCommand.Print, acceleratorOf(SHORTCUTS.print)),
      { type: 'separator' },
      commandItem('menu.file.close', MenuCommand.CloseFile, acceleratorOf(SHORTCUTS.closeFile)),
      { role: 'quit', label: t('menu.file.quit') },
    ],
  }
}

function editMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.edit'),
    submenu: [
      { role: 'undo', label: t('menu.edit.undo') },
      { role: 'redo', label: t('menu.edit.redo') },
      { type: 'separator' },
      { role: 'cut', label: t('menu.edit.cut') },
      { role: 'copy', label: t('menu.edit.copy') },
      { role: 'paste', label: t('menu.edit.paste') },
      commandItem(
        'menu.edit.pasteWithoutFormat',
        MenuCommand.PasteWithoutFormat,
        acceleratorOf(SHORTCUTS.pasteWithoutFormat),
      ),
      { role: 'selectAll', label: t('menu.edit.selectAll') },
      { type: 'separator' },
      commandItem('menu.edit.findReplace', MenuCommand.FindReplace, acceleratorOf(SHORTCUTS.findReplace)),
    ],
  }
}

function formatMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.format'),
    submenu: [
      commandItem('menu.format.paragraph', MenuCommand.ParagraphSetup),
      commandItem('menu.format.image', MenuCommand.ImageProperties),
      commandItem('menu.format.columns', MenuCommand.FormatColumns),
    ],
  }
}

function insertMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.insert'),
    submenu: [
      commandItem(
        'menu.insert.pageBreak',
        MenuCommand.InsertPageBreak,
        acceleratorOf(SHORTCUTS.insertPageBreak),
      ),
      {
        // A quebra de seção não se vê no texto, e por isso também se exclui pelo menu.
        label: t('menu.insert.sectionBreak'),
        submenu: [
          commandItem('menu.insert.sectionNextPage', MenuCommand.InsertSectionNextPage),
          commandItem('menu.insert.sectionContinuous', MenuCommand.InsertSectionContinuous),
          commandItem('menu.insert.sectionEvenPage', MenuCommand.InsertSectionEvenPage),
          commandItem('menu.insert.sectionOddPage', MenuCommand.InsertSectionOddPage),
          { type: 'separator' },
          commandItem('menu.insert.columnBreak', MenuCommand.InsertColumnBreak),
          commandItem('menu.insert.deleteSectionBreak', MenuCommand.DeleteSectionBreak),
        ],
      },
      commandItem('menu.insert.specialCharacter', MenuCommand.SpecialCharacter),
      commandItem(
        'menu.insert.equation',
        MenuCommand.InsertEquation,
        acceleratorOf(SHORTCUTS.insertEquation),
      ),
      commandItem('menu.insert.displayEquation', MenuCommand.InsertDisplayEquation),
      { type: 'separator' },
      commandItem(
        'menu.insert.bookmark',
        MenuCommand.InsertBookmark,
        acceleratorOf(SHORTCUTS.insertBookmark),
      ),
      commandItem(
        'menu.insert.footnote',
        MenuCommand.InsertFootnote,
        acceleratorOf(SHORTCUTS.insertFootnote),
      ),
      commandItem('menu.insert.endnote', MenuCommand.InsertEndnote, acceleratorOf(SHORTCUTS.insertEndnote)),
      commandItem('menu.insert.comment', MenuCommand.InsertComment, acceleratorOf(SHORTCUTS.insertComment)),
      commandItem('menu.insert.nextComment', MenuCommand.NextComment),
      commandItem('menu.insert.previousComment', MenuCommand.PreviousComment),
    ],
  }
}

function referencesMenu(): MenuItemConstructorOptions {
  return {
    // O marcador fica em "Inserir", como no Word.
    label: t('menu.references'),
    submenu: [
      commandItem('menu.references.tableOfContents', MenuCommand.InsertTableOfContents),
      commandItem('menu.references.updateTableOfContents', MenuCommand.UpdateTableOfContents),
      { type: 'separator' },
      commandItem('menu.references.caption', MenuCommand.InsertCaption),
      commandItem('menu.references.crossReference', MenuCommand.InsertCrossReference),
      { type: 'separator' },
      commandItem(
        'menu.references.updateFields',
        MenuCommand.UpdateFields,
        acceleratorOf(SHORTCUTS.updateFields),
      ),
    ],
  }
}

function reviewMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.review'),
    submenu: [
      {
        label: t('revisions.track'),
        type: 'checkbox',
        checked: trackChangesOn,
        accelerator: acceleratorOf(SHORTCUTS.trackChanges),
        toolTip: t('revisions.trackHint'),
        click: () => {
          dispatch(MenuCommand.ToggleTrackChanges)
          // Quem decide é o documento; sem resposta (somente leitura), a marca volta.
          void refreshMenu()
        },
      },
      {
        label: t('revisions.show'),
        submenu: [
          revisionViewItem(RevisionView.All, 'revisions.show.all', MenuCommand.ShowAllMarkup),
          revisionViewItem(RevisionView.Simple, 'revisions.show.simple', MenuCommand.ShowSimpleMarkup),
          revisionViewItem(RevisionView.None, 'revisions.show.none', MenuCommand.ShowNoMarkup),
          revisionViewItem(RevisionView.Original, 'revisions.show.original', MenuCommand.ShowOriginal),
        ],
      },
      { type: 'separator' },
      commandItem('revisions.accept', MenuCommand.AcceptChange),
      commandItem('revisions.reject', MenuCommand.RejectChange),
      { type: 'separator' },
      commandItem('revisions.acceptAll', MenuCommand.AcceptAllChanges),
      commandItem('revisions.rejectAll', MenuCommand.RejectAllChanges),
      { type: 'separator' },
      commandItem('revisions.next', MenuCommand.NextChange),
      commandItem('revisions.previous', MenuCommand.PreviousChange),
    ],
  }
}

function toolsMenu(preferences: EditorPreferences): MenuItemConstructorOptions {
  return {
    label: t('menu.tools'),
    submenu: [
      preferenceToggle('menu.tools.spellcheck', 'spellcheck', preferences),
      preferenceToggle('menu.tools.typography', 'typography', preferences),
      { type: 'separator' },
      commandItem('menu.tools.wordCount', MenuCommand.WordCount, acceleratorOf(SHORTCUTS.wordCount)),
      commandItem('menu.tools.authorName', MenuCommand.AuthorName),
    ],
  }
}

function helpMenu(): MenuItemConstructorOptions {
  return {
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
  }
}

export async function refreshMenu(): Promise<void> {
  Menu.setApplicationMenu(Menu.buildFromTemplate(await buildTemplate()))
}
