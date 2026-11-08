import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { COMMANDS, type CommandMeta } from '@shared/commands/registry.js'
import { APP_NAME } from '@shared/constants.js'
import { LANGUAGES, LANGUAGE_NAMES, type MessageKey } from '@shared/i18n/index.js'
import { SHORTCUTS, acceleratorOf } from '@shared/shortcuts.js'
import { TABLE_ACTIONS } from '@shared/table-actions.js'
import { MenuCommand, RevisionView, Theme, type EditorPreferences } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
import { t } from './i18n.js'
import { logFolder } from './log.js'
import { openFolder } from './open-folder.js'
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

function revisionViewItem(view: RevisionView, command: MenuCommand): MenuItemConstructorOptions {
  return {
    label: t(COMMANDS[command].labelKey),
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
      // Two files with the same name in different folders are common on network shares.
      toolTip: file.path,
      click: () => dispatch(MenuCommand.OpenRecent, file.path),
    })),
    { type: 'separator' },
    commandItem(MenuCommand.ClearRecent),
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

/** The accelerators of Electron's `undo` and `redo` roles. */
const UNDO_ACCELERATOR = 'CommandOrControl+Z'
const REDO_ACCELERATOR = process.platform === 'win32' ? 'Control+Y' : 'Shift+CommandOrControl+Z'

/**
 * Undo and redo are not roles because the role undoes only Chromium's history, and the sheet keeps
 * its own; their keys are the roles', which differ by platform.
 */
const ROLE_ACCELERATORS: Partial<Record<MenuCommand, string>> = {
  [MenuCommand.Undo]: UNDO_ACCELERATOR,
  [MenuCommand.Redo]: REDO_ACCELERATOR,
}

function acceleratorFor(meta: CommandMeta): string | undefined {
  return meta.shortcut === undefined ? ROLE_ACCELERATORS[meta.id] : acceleratorOf(SHORTCUTS[meta.shortcut])
}

function commandItem(command: MenuCommand): MenuItemConstructorOptions {
  const meta: CommandMeta = COMMANDS[command]
  const accelerator = acceleratorFor(meta)
  return {
    label: t(meta.labelKey),
    ...(accelerator === undefined ? {} : { accelerator }),
    ...(meta.registerAccelerator === false ? { registerAccelerator: false } : {}),
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

/**
 * The zoom belongs to the sheet; the step belongs to the renderer, which knows what "fit width" is
 * worth.
 */
function zoomItems(preferences: EditorPreferences): MenuItemConstructorOptions[] {
  return [
    commandItem(MenuCommand.ZoomReset),
    commandItem(MenuCommand.ZoomIn),
    commandItem(MenuCommand.ZoomOut),
    {
      label: t(COMMANDS[MenuCommand.ZoomFitWidth].labelKey),
      type: 'checkbox',
      checked: preferences.zoomFit,
      click: () => dispatch(MenuCommand.ZoomFitWidth),
    },
  ]
}

/** Items are not greyed out outside a table: main does not know where the cursor is. */
function buildTableSubmenu(): MenuItemConstructorOptions[] {
  const items: MenuItemConstructorOptions[] = []
  let group = TABLE_ACTIONS[0]?.group

  for (const action of TABLE_ACTIONS) {
    if (action.group !== group) items.push({ type: 'separator' })
    group = action.group

    items.push(commandItem(action.id))
  }

  return items
}

/** Only items that work. */
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
      // Each language written in itself, outside the catalog: someone looking for "English" would
      // not find "Inglês".
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
      commandItem(MenuCommand.NewDocument),
      commandItem(MenuCommand.NewSpreadsheet),
      commandItem(MenuCommand.NewFromTemplate),
      { type: 'separator' },
      commandItem(MenuCommand.Open),
      { label: t('menu.file.openRecent'), submenu: recent },
      { type: 'separator' },
      commandItem(MenuCommand.Save),
      commandItem(MenuCommand.SaveAs),
      { type: 'separator' },
      commandItem(MenuCommand.PageSetup),
      commandItem(MenuCommand.PrintPreview),
      commandItem(MenuCommand.ExportPdf),
      {
        label: t('menu.file.exportAs'),
        submenu: [
          commandItem(MenuCommand.ExportWord),
          commandItem(MenuCommand.ExportOdt),
          commandItem(MenuCommand.ExportHtml),
          commandItem(MenuCommand.ExportMarkdown),
        ],
      },
      commandItem(MenuCommand.DocumentProperties),
      commandItem(MenuCommand.InspectDocument),
      commandItem(MenuCommand.Print),
      { type: 'separator' },
      commandItem(MenuCommand.CloseFile),
      { role: 'quit', label: t('menu.file.quit') },
    ],
  }
}

function editMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.edit'),
    submenu: [
      commandItem(MenuCommand.Undo),
      commandItem(MenuCommand.Redo),
      { type: 'separator' },
      { role: 'cut', label: t('menu.edit.cut') },
      { role: 'copy', label: t('menu.edit.copy') },
      { role: 'paste', label: t('menu.edit.paste') },
      commandItem(MenuCommand.PasteWithoutFormat),
      { role: 'selectAll', label: t('menu.edit.selectAll') },
      { type: 'separator' },
      commandItem(MenuCommand.FindReplace),
    ],
  }
}

function formatMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.format'),
    submenu: [
      commandItem(MenuCommand.ParagraphSetup),
      commandItem(MenuCommand.ImageProperties),
      commandItem(MenuCommand.FormatColumns),
    ],
  }
}

function insertMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.insert'),
    submenu: [
      commandItem(MenuCommand.InsertPageBreak),
      {
        // A section break is invisible in the text, so it can also be deleted from the menu.
        label: t('menu.insert.sectionBreak'),
        submenu: [
          commandItem(MenuCommand.InsertSectionNextPage),
          commandItem(MenuCommand.InsertSectionContinuous),
          commandItem(MenuCommand.InsertSectionEvenPage),
          commandItem(MenuCommand.InsertSectionOddPage),
          { type: 'separator' },
          commandItem(MenuCommand.InsertColumnBreak),
          commandItem(MenuCommand.DeleteSectionBreak),
        ],
      },
      commandItem(MenuCommand.SpecialCharacter),
      commandItem(MenuCommand.InsertEquation),
      commandItem(MenuCommand.InsertDisplayEquation),
      commandItem(MenuCommand.InsertPropertyField),
      { type: 'separator' },
      commandItem(MenuCommand.InsertBookmark),
      commandItem(MenuCommand.InsertFootnote),
      commandItem(MenuCommand.InsertEndnote),
      commandItem(MenuCommand.InsertComment),
      commandItem(MenuCommand.NextComment),
      commandItem(MenuCommand.PreviousComment),
    ],
  }
}

function referencesMenu(): MenuItemConstructorOptions {
  return {
    // The bookmark lives under "Insert", as in Word.
    label: t('menu.references'),
    submenu: [
      commandItem(MenuCommand.InsertTableOfContents),
      commandItem(MenuCommand.UpdateTableOfContents),
      { type: 'separator' },
      commandItem(MenuCommand.InsertCaption),
      commandItem(MenuCommand.InsertCrossReference),
      { type: 'separator' },
      commandItem(MenuCommand.UpdateFields),
    ],
  }
}

function reviewMenu(): MenuItemConstructorOptions {
  return {
    label: t('menu.review'),
    submenu: [
      {
        ...commandItem(MenuCommand.ToggleTrackChanges),
        type: 'checkbox',
        checked: trackChangesOn,
        toolTip: t('revisions.trackHint'),
        click: () => {
          dispatch(MenuCommand.ToggleTrackChanges)
          // The document decides; without an answer (read-only), the check mark reverts.
          void refreshMenu()
        },
      },
      {
        label: t('revisions.show'),
        submenu: [
          revisionViewItem(RevisionView.All, MenuCommand.ShowAllMarkup),
          revisionViewItem(RevisionView.Simple, MenuCommand.ShowSimpleMarkup),
          revisionViewItem(RevisionView.None, MenuCommand.ShowNoMarkup),
          revisionViewItem(RevisionView.Original, MenuCommand.ShowOriginal),
        ],
      },
      { type: 'separator' },
      commandItem(MenuCommand.AcceptChange),
      commandItem(MenuCommand.RejectChange),
      { type: 'separator' },
      commandItem(MenuCommand.AcceptAllChanges),
      commandItem(MenuCommand.RejectAllChanges),
      { type: 'separator' },
      commandItem(MenuCommand.NextChange),
      commandItem(MenuCommand.PreviousChange),
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
      commandItem(MenuCommand.WordCount),
      commandItem(MenuCommand.AuthorName),
      { type: 'separator' },
      commandItem(MenuCommand.AiAssistant),
      commandItem(MenuCommand.AiProviders),
      { type: 'separator' },
      commandItem(MenuCommand.Options),
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
      { label: t('menu.help.openLogs'), click: () => void openLogFolder() },
    ],
  }
}

async function openLogFolder(): Promise<void> {
  const folder = logFolder()
  if (folder !== null) await openFolder(folder)
}

export async function refreshMenu(): Promise<void> {
  Menu.setApplicationMenu(Menu.buildFromTemplate(await buildTemplate()))
}
