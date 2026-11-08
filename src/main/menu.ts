import { app, BrowserWindow, Menu, type MenuItemConstructorOptions } from 'electron'
import { COMMANDS, type CommandMeta } from '@shared/commands/registry.js'
import { APP_NAME } from '@shared/constants.js'
import { LANGUAGES, LANGUAGE_NAMES, type MessageKey } from '@shared/i18n/index.js'
import { SHORTCUTS, acceleratorOf } from '@shared/shortcuts.js'
import { TABLE_ACTIONS } from '@shared/table-actions.js'
import type { MenuSnapshot } from '@shared/ipc.js'
import { MenuCommand, Theme, type EditorPreferences } from '@shared/types.js'
import { showAboutDialog } from './dialogs.js'
import { listRecentFiles } from './fs/recent.js'
import { t } from './i18n.js'
import { logFolder } from './log.js'
import { openFolder } from './open-folder.js'
import { editorPreferences, updatePreferences } from './preferences.js'
import { devServerUrl, menuSnapshotOf, sendMenuCommand } from './window.js'

const isMac = process.platform === 'darwin'

/** A command's item as the active tab shows it. */
type ItemOf = (command: MenuCommand) => MenuItemConstructorOptions

function itemsFor(snapshot: MenuSnapshot): ItemOf {
  return (command) => {
    const item = commandItem(command)
    return snapshot.disabled.includes(command) ? { ...item, enabled: false } : item
  }
}

function radioItem(command: MenuCommand, snapshot: MenuSnapshot): MenuItemConstructorOptions {
  return {
    ...itemsFor(snapshot)(command),
    type: 'radio',
    checked: snapshot.checked.includes(command),
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

async function buildRecentSubmenu(item: ItemOf): Promise<MenuItemConstructorOptions[]> {
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
    item(MenuCommand.ClearRecent),
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
function zoomItems(preferences: EditorPreferences, item: ItemOf): MenuItemConstructorOptions[] {
  return [
    item(MenuCommand.ZoomReset),
    item(MenuCommand.ZoomIn),
    item(MenuCommand.ZoomOut),
    {
      label: t(COMMANDS[MenuCommand.ZoomFitWidth].labelKey),
      type: 'checkbox',
      checked: preferences.zoomFit,
      click: () => dispatch(MenuCommand.ZoomFitWidth),
    },
  ]
}

/** Items are not greyed out outside a table: main does not know where the cursor is. */
function buildTableSubmenu(item: ItemOf): MenuItemConstructorOptions[] {
  const items: MenuItemConstructorOptions[] = []
  let group = TABLE_ACTIONS[0]?.group

  for (const action of TABLE_ACTIONS) {
    if (action.group !== group) items.push({ type: 'separator' })
    group = action.group

    items.push(item(action.id))
  }

  return items
}

/** Only items that work; the state is the focused window's active tab. */
async function buildTemplate(snapshot: MenuSnapshot): Promise<MenuItemConstructorOptions[]> {
  const preferences = editorPreferences()
  const item = itemsFor(snapshot)

  return [
    ...macAppMenu(),
    fileMenu(await buildRecentSubmenu(item), item),
    editMenu(item),
    formatMenu(item),
    { label: t('menu.table'), submenu: buildTableSubmenu(item) },
    insertMenu(item),
    referencesMenu(item),
    reviewMenu(snapshot),
    { label: t('menu.view'), submenu: viewSubmenu(preferences, item) },
    toolsMenu(preferences, item),
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

function viewSubmenu(preferences: EditorPreferences, item: ItemOf): MenuItemConstructorOptions[] {
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
    ...zoomItems(preferences, item),
    { type: 'separator' },
    item(MenuCommand.NextTab),
    item(MenuCommand.PreviousTab),
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

function fileMenu(recent: MenuItemConstructorOptions[], item: ItemOf): MenuItemConstructorOptions {
  return {
    label: t('menu.file'),
    submenu: [
      item(MenuCommand.NewDocument),
      item(MenuCommand.NewSpreadsheet),
      item(MenuCommand.NewFromTemplate),
      { type: 'separator' },
      item(MenuCommand.Open),
      { label: t('menu.file.openRecent'), submenu: recent },
      item(MenuCommand.ShowHome),
      { type: 'separator' },
      item(MenuCommand.Save),
      item(MenuCommand.SaveAs),
      { type: 'separator' },
      item(MenuCommand.PageSetup),
      item(MenuCommand.PrintPreview),
      item(MenuCommand.ExportPdf),
      {
        label: t('menu.file.exportAs'),
        submenu: [
          item(MenuCommand.ExportWord),
          item(MenuCommand.ExportOdt),
          item(MenuCommand.ExportHtml),
          item(MenuCommand.ExportMarkdown),
        ],
      },
      item(MenuCommand.DocumentProperties),
      item(MenuCommand.InspectDocument),
      item(MenuCommand.Print),
      { type: 'separator' },
      item(MenuCommand.CloseFile),
      item(MenuCommand.CloseAllTabs),
      { role: 'quit', label: t('menu.file.quit') },
    ],
  }
}

function editMenu(item: ItemOf): MenuItemConstructorOptions {
  return {
    label: t('menu.edit'),
    submenu: [
      item(MenuCommand.Undo),
      item(MenuCommand.Redo),
      { type: 'separator' },
      { role: 'cut', label: t('menu.edit.cut') },
      { role: 'copy', label: t('menu.edit.copy') },
      { role: 'paste', label: t('menu.edit.paste') },
      item(MenuCommand.PasteWithoutFormat),
      { role: 'selectAll', label: t('menu.edit.selectAll') },
      { type: 'separator' },
      item(MenuCommand.FindReplace),
    ],
  }
}

function formatMenu(item: ItemOf): MenuItemConstructorOptions {
  return {
    label: t('menu.format'),
    submenu: [
      item(MenuCommand.ParagraphSetup),
      item(MenuCommand.ImageProperties),
      item(MenuCommand.FormatColumns),
    ],
  }
}

function insertMenu(item: ItemOf): MenuItemConstructorOptions {
  return {
    label: t('menu.insert'),
    submenu: [
      item(MenuCommand.InsertPageBreak),
      {
        // A section break is invisible in the text, so it can also be deleted from the menu.
        label: t('menu.insert.sectionBreak'),
        submenu: [
          item(MenuCommand.InsertSectionNextPage),
          item(MenuCommand.InsertSectionContinuous),
          item(MenuCommand.InsertSectionEvenPage),
          item(MenuCommand.InsertSectionOddPage),
          { type: 'separator' },
          item(MenuCommand.InsertColumnBreak),
          item(MenuCommand.DeleteSectionBreak),
        ],
      },
      item(MenuCommand.SpecialCharacter),
      item(MenuCommand.InsertEquation),
      item(MenuCommand.InsertDisplayEquation),
      item(MenuCommand.InsertPropertyField),
      { type: 'separator' },
      item(MenuCommand.InsertBookmark),
      item(MenuCommand.InsertFootnote),
      item(MenuCommand.InsertEndnote),
      item(MenuCommand.InsertComment),
      item(MenuCommand.NextComment),
      item(MenuCommand.PreviousComment),
    ],
  }
}

function referencesMenu(item: ItemOf): MenuItemConstructorOptions {
  return {
    // The bookmark lives under "Insert", as in Word.
    label: t('menu.references'),
    submenu: [
      item(MenuCommand.InsertTableOfContents),
      item(MenuCommand.UpdateTableOfContents),
      { type: 'separator' },
      item(MenuCommand.InsertCaption),
      item(MenuCommand.InsertCrossReference),
      { type: 'separator' },
      item(MenuCommand.UpdateFields),
    ],
  }
}

function reviewMenu(snapshot: MenuSnapshot): MenuItemConstructorOptions {
  const item = itemsFor(snapshot)
  return {
    label: t('menu.review'),
    submenu: [
      {
        ...item(MenuCommand.ToggleTrackChanges),
        type: 'checkbox',
        checked: snapshot.checked.includes(MenuCommand.ToggleTrackChanges),
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
          radioItem(MenuCommand.ShowAllMarkup, snapshot),
          radioItem(MenuCommand.ShowSimpleMarkup, snapshot),
          radioItem(MenuCommand.ShowNoMarkup, snapshot),
          radioItem(MenuCommand.ShowOriginal, snapshot),
        ],
      },
      { type: 'separator' },
      item(MenuCommand.AcceptChange),
      item(MenuCommand.RejectChange),
      { type: 'separator' },
      item(MenuCommand.AcceptAllChanges),
      item(MenuCommand.RejectAllChanges),
      { type: 'separator' },
      item(MenuCommand.NextChange),
      item(MenuCommand.PreviousChange),
    ],
  }
}

function toolsMenu(preferences: EditorPreferences, item: ItemOf): MenuItemConstructorOptions {
  return {
    label: t('menu.tools'),
    submenu: [
      preferenceToggle('menu.tools.spellcheck', 'spellcheck', preferences),
      preferenceToggle('menu.tools.typography', 'typography', preferences),
      { type: 'separator' },
      item(MenuCommand.WordCount),
      item(MenuCommand.AuthorName),
      { type: 'separator' },
      item(MenuCommand.AiAssistant),
      item(MenuCommand.AiProviders),
      { type: 'separator' },
      item(MenuCommand.Options),
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
  Menu.setApplicationMenu(Menu.buildFromTemplate(await buildTemplate(menuSnapshotOf(focusedWindow()))))
}
