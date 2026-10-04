import { useEffect } from 'react'
import { MenuCommand } from '@shared/types.js'
import { buildWindowTitle } from '@services/file/formats.js'
import type { WorkbookModel } from '@services/spreadsheet/model.js'
import { ErrorBanner } from '../components/ErrorBanner.js'
import { InventoryBanner } from '../components/InventoryBanner.js'
import { ReadOnlyBanner } from '../components/ReadOnlyBanner.js'
import { RecoveryBanner } from '../components/RecoveryBanner.js'
import { StatusBar } from '../components/StatusBar.js'
import { TemplateGallery } from '../components/TemplateGallery.js'
import { DocumentEditor } from '../document/DocumentEditor.js'
import { asEditorCommand, emitEditorCommand } from '../document/editor-commands.js'
import { HomePage } from '../pages/HomePage.js'
import { SheetTabs } from '../spreadsheet/SheetTabs.js'
import { SpreadsheetEditor } from '../spreadsheet/SpreadsheetEditor.js'
import { usePreferences, watchPreferences } from '../state/preferences.js'
import { useReadingMode } from '../state/reading.js'
import { useTheme } from '../state/theme.js'
import { runZoomCommand } from '../state/zoom.js'
import { revisionViewOfCommand, setRevisionView, useRevisionView } from '../state/revision-view.js'
import { useWorkspace } from '../state/workspace.js'
import { t } from '../i18n.js'

/** O teto de trabalho que uma queda pode custar; menos pesaria em documento grande. */
const AUTOSAVE_INTERVAL_MS = 8_000

async function runMenuCommand(command: MenuCommand, path: string | undefined): Promise<void> {
  const workspace = useWorkspace.getState()

  // Os comandos do editor, reconhecidos pelo nome: o App não tem referência a ele.
  const editorCommand = asEditorCommand(command)
  if (editorCommand !== null) return emitEditorCommand(editorCommand)
  const revisionView = revisionViewOfCommand(command)
  if (revisionView !== null) return setRevisionView(revisionView)

  await MENU_ACTIONS[command]?.(workspace, path)
}

type Workspace = ReturnType<typeof useWorkspace.getState>
type MenuAction = (workspace: Workspace, path: string | undefined) => Promise<unknown> | unknown

const MENU_ACTIONS: Partial<Record<MenuCommand, MenuAction>> = {
  [MenuCommand.NewDocument]: (workspace) => workspace.newDocument(),
  [MenuCommand.NewSpreadsheet]: (workspace) => workspace.newSpreadsheet(),
  [MenuCommand.NewFromTemplate]: (workspace) => workspace.setTemplateGallery(true),
  [MenuCommand.Open]: (workspace) => workspace.openViaDialog(),
  [MenuCommand.OpenRecent]: (workspace, path) =>
    path === undefined ? undefined : workspace.openRecent(path),
  [MenuCommand.ClearRecent]: (workspace) => workspace.clearRecents(),
  [MenuCommand.Save]: (workspace) => workspace.save(),
  [MenuCommand.SaveAs]: (workspace) => workspace.saveAs(),
  [MenuCommand.CloseFile]: (workspace) => workspace.closeFile(),
  // Só fecha se a gravação der certo.
  [MenuCommand.SaveAndExit]: async (workspace) => {
    if (await workspace.save()) await window.api.window.close({})
  },
  [MenuCommand.ExportPdf]: (workspace) => workspace.exportPdf(),
  [MenuCommand.ExportHtml]: (workspace) => workspace.exportDocument('html'),
  [MenuCommand.ExportMarkdown]: (workspace) => workspace.exportDocument('markdown'),
  [MenuCommand.ExportOdt]: (workspace) => workspace.exportDocument('odt'),
  [MenuCommand.Print]: (workspace) => workspace.print(),
  [MenuCommand.PrintPreview]: (workspace) => workspace.printPreview(),
  [MenuCommand.ZoomIn]: () => runZoomCommand(MenuCommand.ZoomIn),
  [MenuCommand.ZoomOut]: () => runZoomCommand(MenuCommand.ZoomOut),
  [MenuCommand.ZoomReset]: () => runZoomCommand(MenuCommand.ZoomReset),
  [MenuCommand.ZoomFitWidth]: () => runZoomCommand(MenuCommand.ZoomFitWidth),
}

export function App(): React.JSX.Element {
  const hasFile = useWorkspace((state) => state.file !== null)
  const templateGallery = useWorkspace((state) => state.templateGallery)
  // O editor é recarregado a cada documento, sem estado residual.
  const generation = useWorkspace((state) => state.generation)
  const workbook = useWorkspace((state) => state.workbook)
  const reading = useReadingMode()
  const showStatusBar = usePreferences((state) => state.preferences.showStatusBar)

  useWorkspaceLifecycle()
  useWindowStateSync()

  // Azul de documento, verde de planilha: ver `--accent` no CSS.
  return (
    <div
      className={['app', workbook === null ? '' : 'app--spreadsheet', reading ? 'app--reading' : '']
        .filter((name) => name !== '')
        .join(' ')}
    >
      <ErrorBanner />
      <RecoveryBanner />
      <ReadOnlyBanner />
      <InventoryBanner />
      <div className="app__body">
        {workbook !== null ? (
          <WorkbookView workbook={workbook} generation={generation} />
        ) : hasFile ? (
          <DocumentEditor key={generation} />
        ) : (
          <HomePage />
        )}
      </div>
      {/* Contagem de palavras e número de páginas são ferramentas de quem escreve. */}
      {hasFile && !reading && showStatusBar && <StatusBar />}
      {templateGallery && <TemplateGallery />}
    </div>
  )
}

/** Recentes, recuperação, preferências, tema, autosave e o menu do main. */
function useWorkspaceLifecycle(): void {
  useEffect(() => {
    void useWorkspace.getState().refreshRecents()
    void useWorkspace.getState().checkRecovery()
  }, [])

  // A cópia das preferências que a tela desenha; o dono é o main.
  useEffect(() => watchPreferences(), [])

  // Inclusive quando quem mudou o tema foi o sistema operacional.
  useTheme()

  useEffect(() => {
    // Por relógio, e não por tecla: o autosave serializa o documento inteiro.
    const timer = setInterval(() => {
      void useWorkspace.getState().autosave()
    }, AUTOSAVE_INTERVAL_MS)
    return () => clearInterval(timer)
  }, [])

  useEffect(() => {
    let commands = Promise.resolve()
    const unsubscribe = window.api.menu.onCommand(({ command, path }) => {
      commands = commands.then(() => runMenuCommand(command, path)).catch(console.error)
    })
    // Assina antes de pedir ao main os arquivos escolhidos no Explorer.
    void window.api.window.ready({})
    return unsubscribe
  }, [])
}

/** O título da janela e o estado que o menu do main mostra. */
function useWindowStateSync(): void {
  useEffect(() => {
    // Só quando o título ou o "não salvo" mudam, e não a cada tecla.
    let lastTitle = ''
    let lastDirty: boolean | null = null
    let lastTracking: boolean | null = null
    let lastView: string | null = null

    const sync = (): void => {
      const state = useWorkspace.getState()
      const untitled = t('shell.file.untitled')
      const title = state.file?.name ?? untitled
      // É do documento: com planilha aberta, desligado.
      const trackChanges = state.workbook === null && state.trackChanges === true
      const revisionView = useRevisionView.getState().view
      if (
        title === lastTitle &&
        state.isDirty === lastDirty &&
        trackChanges === lastTracking &&
        revisionView === lastView
      )
        return

      lastTitle = title
      lastDirty = state.isDirty
      lastTracking = trackChanges
      lastView = revisionView
      void window.api.window.setState({ title, isDirty: state.isDirty, trackChanges, revisionView })
      document.title = buildWindowTitle(state.file?.name ?? null, state.isDirty, 'Librevia', untitled)
    }

    sync()
    const unsubscribeView = useRevisionView.subscribe(sync)
    const unsubscribe = useWorkspace.subscribe(sync)
    return () => {
      unsubscribeView()
      unsubscribe()
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
  // Uma ação por seletor: um objeto novo a cada chamada faria o React entrar em laço.
  const selectSheet = useWorkspace((state) => state.selectSheet)
  const addSheet = useWorkspace((state) => state.addSheet)
  const renameSheet = useWorkspace((state) => state.renameSheet)
  const removeSheet = useWorkspace((state) => state.removeSheet)
  const readOnly = useWorkspace((state) => state.readOnly)
  return (
    <div className="workbook">
      <SpreadsheetEditor
        key={`${generation}-${workbook.activeSheet}`}
        sheet={workbook.sheets[workbook.activeSheet]!}
        onChange={updateSheet}
        onStructure={changeStructure}
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
  )
}
