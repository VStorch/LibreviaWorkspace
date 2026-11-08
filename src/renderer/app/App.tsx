import { useEffect, useState } from 'react'
import { DocumentKind, EditCommand, MenuCommand } from '@shared/types.js'
import type { WorkbookModel } from '@services/spreadsheet/model.js'
import { ErrorBanner } from '../components/ErrorBanner.js'
import { OptionsDialog } from '../components/OptionsDialog.js'
import { PrivacyInspector } from '../components/PrivacyInspector.js'
import { InventoryBanner } from '../components/InventoryBanner.js'
import { ReadOnlyBanner } from '../components/ReadOnlyBanner.js'
import { RecoveryBanner } from '../components/RecoveryBanner.js'
import { StatusBar } from '../components/StatusBar.js'
import { answerMcpRequests } from '../document/mcp-document.js'
import { watchMcpClients } from '../state/mcp-clients.js'
import { TemplateGallery } from '../components/TemplateGallery.js'
import { DocumentEditor } from '../document/DocumentEditor.js'
import { EditorCommand } from '../document/editor-commands.js'
import { useEditorHandlers } from '../commands/useEditorHandlers.js'
import { executeCommand, registerHandlers, type CommandHandlers } from '../commands/execute.js'
import { HomePage } from '../pages/HomePage.js'
import { PdfView } from '../pdf/PdfView.js'
import { SheetTabs } from '../spreadsheet/SheetTabs.js'
import { SpreadsheetEditor } from '../spreadsheet/SpreadsheetEditor.js'
import { SheetAiPanel } from '../spreadsheet/SheetAiPanel.js'
import { SheetMcpProposal } from '../spreadsheet/SheetMcpProposal.js'
import { useOfferSheetToMcp } from '../spreadsheet/mcp-sheet.js'
import { useSheetAiPane } from '../spreadsheet/useSheetAiPane.js'
import { sheetOwnsUndo, useSheetUndo } from '../spreadsheet/useSheetUndo.js'
import { AiProvidersDialog } from '../document/AiProvidersDialog.js'
import { singleCell, type Range } from '@services/spreadsheet/edit.js'
import { WorkbookPropertiesDialog } from '../spreadsheet/WorkbookPropertiesDialog.js'
import { usePreferences, watchPreferences } from '../state/preferences.js'
import { useReadingMode } from '../state/reading.js'
import { setOptionsDialog, useOptionsDialog } from '../state/options-dialog.js'
import { useTheme } from '../state/theme.js'
import { runZoomCommand } from '../state/zoom.js'
import { REVISION_VIEW_HANDLERS } from '../state/revision-view.js'
import {
  activateNextTab,
  activatePreviousTab,
  activeWorkspaceStore,
  autosaveAllTabs,
  showHome,
  useTabs,
} from '../state/tabs.js'
import { useWorkspace } from '../state/workspace.js'
import { refreshRecents } from '../state/recents.js'
import {
  checkRecovery,
  clearRecents,
  closeActiveTab,
  closeAllTabs,
  newDocument,
  newSpreadsheet,
  openRecent,
  openViaDialog,
  saveAllTabs,
  setTemplateGallery,
  useShell,
} from '../state/shell.js'
import { TabBar } from '../components/TabBar.js'
import { useWindowStateSync } from './window-state.js'
import { useTabKeys } from './tab-keys.js'

/** The most work a crash can cost; less would weigh on large documents. */
const AUTOSAVE_INTERVAL_MS = 8_000
const WORKBOOK_COMMANDS = [EditorCommand.DocumentProperties]

type Workspace = ReturnType<typeof useWorkspace.getState>

type HistoryCommand = typeof EditCommand.Undo | typeof EditCommand.Redo

const HISTORY_INPUT: Record<HistoryCommand, string> = {
  [EditCommand.Undo]: 'historyUndo',
  [EditCommand.Redo]: 'historyRedo',
}

/**
 * The editors hear the menu as the browser's own undo, a `beforeinput` they take; Chromium's redo
 * would stay greyed out, since the editor kept its undo away from it.
 */
function undoInEditor(command: HistoryCommand): boolean {
  const focused = document.activeElement
  if (!(focused instanceof HTMLElement) || !focused.isContentEditable) return false
  const event = new InputEvent('beforeinput', {
    inputType: HISTORY_INPUT[command],
    bubbles: true,
    cancelable: true,
  })
  return !focused.dispatchEvent(event)
}

/** From the menu: the sheet's history outside its text fields; anywhere else, the focused one's. */
function runHistory(workspace: Workspace, command: HistoryCommand): unknown {
  if (workspace.workbook !== null && sheetOwnsUndo(document.activeElement)) {
    return command === EditCommand.Undo ? workspace.undoSheet() : workspace.redoSheet()
  }
  if (undoInEditor(command)) return undefined
  return window.api.edit.run({ command })
}

type MenuAction = (workspace: Workspace, path: string | undefined) => Promise<unknown> | unknown

const MENU_ACTIONS: Partial<Record<MenuCommand, MenuAction>> = {
  [MenuCommand.Undo]: (workspace) => runHistory(workspace, EditCommand.Undo),
  [MenuCommand.Redo]: (workspace) => runHistory(workspace, EditCommand.Redo),
  [MenuCommand.NewDocument]: () => newDocument(),
  [MenuCommand.NewSpreadsheet]: () => newSpreadsheet(),
  [MenuCommand.NewFromTemplate]: () => setTemplateGallery(true),
  [MenuCommand.Open]: () => openViaDialog(),
  [MenuCommand.OpenRecent]: (_workspace, path) => (path === undefined ? undefined : openRecent(path)),
  [MenuCommand.ClearRecent]: () => clearRecents(),
  [MenuCommand.Save]: (workspace) => workspace.save(),
  [MenuCommand.SaveAs]: (workspace) => workspace.saveAs(),
  [MenuCommand.CloseFile]: () => closeActiveTab(),
  [MenuCommand.CloseAllTabs]: () => closeAllTabs(),
  [MenuCommand.ShowHome]: () => showHome(),
  [MenuCommand.NextTab]: () => activateNextTab(),
  [MenuCommand.PreviousTab]: () => activatePreviousTab(),
  // Only closes if every save succeeds.
  [MenuCommand.SaveAndExit]: async () => {
    if (await saveAllTabs()) await window.api.window.close({})
  },
  [MenuCommand.InspectDocument]: (workspace) => workspace.setInspector(workspace.file !== null),
  [MenuCommand.ExportPdf]: (workspace) => workspace.exportPdf(),
  [MenuCommand.ExportHtml]: (workspace) => workspace.exportDocument('html'),
  [MenuCommand.ExportMarkdown]: (workspace) => workspace.exportDocument('markdown'),
  [MenuCommand.ExportOdt]: (workspace) => workspace.exportDocument('odt'),
  [MenuCommand.ExportWord]: (workspace) => workspace.saveAs({ wordOnly: true }),
  [MenuCommand.Print]: (workspace) => workspace.print(),
  [MenuCommand.PrintPreview]: (workspace) => workspace.printPreview(),
  [MenuCommand.ZoomIn]: () => runZoomCommand(MenuCommand.ZoomIn),
  [MenuCommand.ZoomOut]: () => runZoomCommand(MenuCommand.ZoomOut),
  [MenuCommand.ZoomReset]: () => runZoomCommand(MenuCommand.ZoomReset),
  [MenuCommand.ZoomFitWidth]: () => runZoomCommand(MenuCommand.ZoomFitWidth),
  [MenuCommand.Options]: () => setOptionsDialog(true),
}

function shellHandlers(): CommandHandlers {
  const entries = Object.entries(MENU_ACTIONS) as [MenuCommand, MenuAction][]
  return Object.fromEntries(
    entries.map(([id, action]) => [
      id,
      {
        run: async (arg?: string) => {
          await action(activeWorkspaceStore().getState(), arg)
        },
      },
    ]),
  )
}

export function App(): React.JSX.Element {
  const hasFile = useWorkspace((state) => state.file !== null)
  const viewedPdf = useWorkspace((state) => (state.file?.kind === DocumentKind.Pdf ? state.file.path : null))
  const templateGallery = useShell((state) => state.templateGallery)
  const inspector = useWorkspace((state) => state.inspector && state.file !== null)
  // The editor is remounted for each document and each tab, with no leftover state.
  const tabId = useTabs((tabs) => tabs.activeId)
  const generation = useWorkspace((state) => state.generation)
  const workbook = useWorkspace((state) => state.workbook)
  const reading = useReadingMode()
  const showStatusBar = usePreferences((state) => state.preferences.showStatusBar)
  const options = useOptionsDialog((state) => state.open)

  useWorkspaceLifecycle()
  useWindowStateSync()
  useTabKeys()

  // Blue for documents, green for spreadsheets: see `--accent` in the CSS.
  return (
    <div
      className={['app', workbook === null ? '' : 'app--spreadsheet', reading ? 'app--reading' : '']
        .filter((name) => name !== '')
        .join(' ')}
    >
      <TabBar />
      <ErrorBanner />
      <RecoveryBanner />
      <ReadOnlyBanner />
      <InventoryBanner />
      <div className="app__body">
        {viewedPdf !== null ? (
          <PdfView key={`${tabId}:${generation}`} path={viewedPdf} />
        ) : workbook !== null ? (
          <WorkbookView key={tabId} workbook={workbook} generation={generation} />
        ) : hasFile ? (
          <DocumentEditor key={`${tabId}:${generation}`} />
        ) : (
          <HomePage />
        )}
      </div>
      {/* Word count and page number are tools for writing. */}
      {hasFile && !reading && showStatusBar && <StatusBar />}
      {templateGallery && <TemplateGallery />}
      {inspector && <PrivacyInspector />}
      {options && <OptionsDialog onClose={() => setOptionsDialog(false)} />}
    </div>
  )
}

/** Recent files, recovery, preferences, theme, autosave and the main menu. */
function useWorkspaceLifecycle(): void {
  useEffect(() => {
    void refreshRecents()
    void checkRecovery()
  }, [])

  // The copy of the preferences the screen draws; main owns them.
  useEffect(() => watchPreferences(), [])

  // The local MCP server's clients: who is connected, and what they read and change.
  useEffect(() => watchMcpClients(), [])
  useEffect(() => answerMcpRequests(), [])

  // Including when the operating system changed the theme.
  useTheme()

  useEffect(() => {
    // By timer, not by key press: autosave serializes the whole document.
    const timer = setInterval(() => {
      void autosaveAllTabs()
    }, AUTOSAVE_INTERVAL_MS)
    return () => clearInterval(timer)
  }, [])

  useEffect(() => {
    const unregisterShell = registerHandlers('shell', shellHandlers())
    const unregisterRevisions = registerHandlers('revisions', REVISION_VIEW_HANDLERS)
    let commands = Promise.resolve()
    const unsubscribe = window.api.menu.onCommand(({ command, path }) => {
      commands = commands.then(() => executeCommand(command, path)).catch(console.error)
    })
    // Subscribes before asking main for files picked in the file manager.
    void window.api.window.ready({})
    return () => {
      unsubscribe()
      unregisterRevisions()
      unregisterShell()
    }
  }, [])
}

function WorkbookView({
  workbook,
  generation,
}: {
  workbook: WorkbookModel
  generation: number
}): React.JSX.Element {
  const updateSheet = useWorkspace((state) => state.updateSheet)
  const changeStructure = useWorkspace((state) => state.changeStructure)
  // One action per selector: a new object on each call would send React into a loop.
  const selectSheet = useWorkspace((state) => state.selectSheet)
  const addSheet = useWorkspace((state) => state.addSheet)
  const renameSheet = useWorkspace((state) => state.renameSheet)
  const removeSheet = useWorkspace((state) => state.removeSheet)
  const undoSheet = useWorkspace((state) => state.undoSheet)
  const redoSheet = useWorkspace((state) => state.redoSheet)
  const readOnly = useWorkspace((state) => state.readOnly)
  const [showProperties, setShowProperties] = useState(false)
  const [range, setRange] = useState<Range>(() => singleCell(0, 0))
  const aiPane = useSheetAiPane()
  useSheetUndo(undoSheet, redoSheet)
  useOfferSheetToMcp(range)
  useEditorHandlers(WORKBOOK_COMMANDS, () => setShowProperties(true))

  return (
    <div className="workbook-shell">
      <div className="workbook">
        <SheetMcpProposal />
        <SpreadsheetEditor
          key={`${generation}-${workbook.activeSheet}`}
          sheet={workbook.sheets[workbook.activeSheet]!}
          onChange={updateSheet}
          onStructure={changeStructure}
          onRange={setRange}
          readOnly={readOnly}
        />
        <SheetTabs
          workbook={workbook}
          onSelect={selectSheet}
          onAdd={addSheet}
          onRename={renameSheet}
          onRemove={removeSheet}
        />
      </div>
      {aiPane.open && (
        <SheetAiPanel
          workbook={workbook}
          range={range}
          readOnly={readOnly}
          configuring={aiPane.configuring}
          onApply={(sheet) => {
            if (!readOnly) updateSheet(sheet)
          }}
          onClose={aiPane.close}
          onConfigure={() => aiPane.configure(true)}
        />
      )}
      {aiPane.configuring && <AiProvidersDialog onClose={() => aiPane.configure(false)} />}
      {showProperties && (
        <WorkbookPropertiesDialog workbook={workbook} onClose={() => setShowProperties(false)} />
      )}
    </div>
  )
}
