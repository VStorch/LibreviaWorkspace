import { useEffect } from 'react'
import { MenuCommand } from '@shared/types.js'
import { buildWindowTitle } from '@services/file/formats.js'
import { ErrorBanner } from '../components/ErrorBanner.js'
import { InventoryBanner } from '../components/InventoryBanner.js'
import { ReadOnlyBanner } from '../components/ReadOnlyBanner.js'
import { RecoveryBanner } from '../components/RecoveryBanner.js'
import { StatusBar } from '../components/StatusBar.js'
import { DocumentEditor } from '../document/DocumentEditor.js'
import { emitEditorCommand } from '../document/editor-commands.js'
import { HomePage } from '../pages/HomePage.js'
import { SheetTabs } from '../spreadsheet/SheetTabs.js'
import { SpreadsheetEditor } from '../spreadsheet/SpreadsheetEditor.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * De quanto em quanto tempo o rascunho é regravado.
 *
 * Oito segundos é o teto de trabalho que uma queda pode custar. Menos que isso
 * começaria a pesar em documento grande, onde cada gravação serializa tudo.
 */
const AUTOSAVE_INTERVAL_MS = 8_000

/** Traduz um comando do menu nativo na ação correspondente. */
async function runMenuCommand(command: MenuCommand, path: string | undefined): Promise<void> {
  const workspace = useWorkspace.getState()

  switch (command) {
    case MenuCommand.NewDocument:
      return workspace.newDocument()
    case MenuCommand.Open:
      return workspace.openViaDialog()
    case MenuCommand.OpenRecent:
      if (path !== undefined) await workspace.openRecent(path)
      return
    case MenuCommand.ClearRecent:
      return workspace.clearRecents()
    case MenuCommand.Save:
      await workspace.save()
      return
    case MenuCommand.SaveAs:
      await workspace.saveAs()
      return
    case MenuCommand.CloseFile:
      return workspace.closeFile()
    case MenuCommand.SaveAndExit: {
      // O usuário escolheu "Salvar" no aviso de saída: só fechamos se a
      // gravação der certo, senão a janela sumiria levando o trabalho junto.
      if (await workspace.save()) await window.api.window.close({})
      return
    }

    case MenuCommand.ExportPdf:
      await workspace.exportPdf()
      return
    case MenuCommand.Print:
      await workspace.print()
      return
    case MenuCommand.PrintPreview:
      return workspace.printPreview()

    // Comandos que pertencem ao editor: o App não tem referência a ele.
    case MenuCommand.FindReplace:
      return emitEditorCommand('find-replace')
    case MenuCommand.PageSetup:
      return emitEditorCommand('page-setup')
    case MenuCommand.InsertPageBreak:
      return emitEditorCommand('insert-page-break')

    case MenuCommand.NewSpreadsheet:
      return useWorkspace.getState().newSpreadsheet()
  }
}

export function App(): React.JSX.Element {
  const hasFile = useWorkspace((state) => state.file !== null)
  // Recarrega o editor por completo a cada documento aberto, em vez de tentar
  // sincronizar conteúdo — elimina estado residual entre um arquivo e outro.
  const generation = useWorkspace((state) => state.generation)
  const workbook = useWorkspace((state) => state.workbook)
  const updateSheet = useWorkspace((state) => state.updateSheet)
  const changeStructure = useWorkspace((state) => state.changeStructure)
  // Uma ação por seletor: devolver um objeto novo a cada chamada faria o
  // zustand ver estado diferente toda renderização, e o React entraria em laço.
  const selectSheet = useWorkspace((state) => state.selectSheet)
  const addSheet = useWorkspace((state) => state.addSheet)
  const renameSheet = useWorkspace((state) => state.renameSheet)
  const removeSheet = useWorkspace((state) => state.removeSheet)
  const readOnly = useWorkspace((state) => state.readOnly)

  useEffect(() => {
    void useWorkspace.getState().refreshRecents()
    void useWorkspace.getState().checkRecovery()
  }, [])

  useEffect(() => {
    // Por relógio, e não por tecla: o autosave serializa o documento inteiro, e
    // fazer isso a cada caractere digitado travaria a digitação num arquivo
    // grande. O intervalo é o teto de trabalho que uma queda pode custar.
    const timer = setInterval(() => {
      void useWorkspace.getState().autosave()
    }, AUTOSAVE_INTERVAL_MS)
    return () => clearInterval(timer)
  }, [])

  useEffect(
    () =>
      window.api.menu.onCommand(({ command, path }) => {
        void runMenuCommand(command, path)
      }),
    [],
  )

  useEffect(() => {
    // O título e o marcador de "não salvo" vivem no main. Só enviamos quando
    // algum dos dois muda de fato — não a cada tecla digitada.
    let lastTitle = ''
    let lastDirty: boolean | null = null

    const sync = (): void => {
      const state = useWorkspace.getState()
      const title = state.file?.name ?? 'Sem título'
      if (title === lastTitle && state.isDirty === lastDirty) return

      lastTitle = title
      lastDirty = state.isDirty
      void window.api.window.setState({ title, isDirty: state.isDirty })
      document.title = buildWindowTitle(state.file?.name ?? null, state.isDirty, 'Librevia')
    }

    sync()
    return useWorkspace.subscribe(sync)
  }, [])

  // A casca inteira muda de cor conforme o que está aberto — azul de
  // documento, verde de planilha. Ver o comentário de `--accent` no CSS.
  return (
    <div className={workbook === null ? 'app' : 'app app--spreadsheet'}>
      <ErrorBanner />
      <RecoveryBanner />
      <ReadOnlyBanner />
      <InventoryBanner />
      <div className="app__body">
        {workbook !== null ? (
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
        ) : hasFile ? (
          <DocumentEditor key={generation} />
        ) : (
          <HomePage />
        )}
      </div>
      {hasFile && <StatusBar />}
    </div>
  )
}
