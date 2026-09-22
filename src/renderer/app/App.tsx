import { useEffect } from 'react'
import { MenuCommand } from '@shared/types.js'
import { buildWindowTitle } from '@services/file/formats.js'
import { ErrorBanner } from '../components/ErrorBanner.js'
import { InventoryBanner } from '../components/InventoryBanner.js'
import { ReadOnlyBanner } from '../components/ReadOnlyBanner.js'
import { RecoveryBanner } from '../components/RecoveryBanner.js'
import { StatusBar } from '../components/StatusBar.js'
import { DocumentEditor } from '../document/DocumentEditor.js'
import { asEditorCommand, emitEditorCommand } from '../document/editor-commands.js'
import { HomePage } from '../pages/HomePage.js'
import { SheetTabs } from '../spreadsheet/SheetTabs.js'
import { SpreadsheetEditor } from '../spreadsheet/SpreadsheetEditor.js'
import { watchPreferences } from '../state/preferences.js'
import { useReadingMode } from '../state/reading.js'
import { useTheme } from '../state/theme.js'
import { useWorkspace } from '../state/workspace.js'
import { t } from '../i18n.js'

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

  // Comandos que pertencem ao editor: o App não tem referência a ele, e os
  // reconhece pelo nome — o mesmo nas duas pontas. Uma tradução caso a caso
  // crescia um `case` por recurso, e as tabelas sozinhas trouxeram doze.
  const editorCommand = asEditorCommand(command)
  if (editorCommand !== null) return emitEditorCommand(editorCommand)

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
  const reading = useReadingMode()

  useEffect(() => {
    void useWorkspace.getState().refreshRecents()
    void useWorkspace.getState().checkRecovery()
  }, [])

  // As preferências de edição moram no main, que é quem liga o corretor na sessão
  // do Chromium. Aqui só se mantém a cópia que a tela desenha.
  useEffect(() => watchPreferences(), [])

  // Escreve `data-theme` na raiz e o mantém em dia — inclusive quando quem
  // mudou foi o sistema operacional, e não o menu.
  useTheme()

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
      const untitled = t('shell.file.untitled')
      const title = state.file?.name ?? untitled
      if (title === lastTitle && state.isDirty === lastDirty) return

      lastTitle = title
      lastDirty = state.isDirty
      void window.api.window.setState({ title, isDirty: state.isDirty })
      document.title = buildWindowTitle(state.file?.name ?? null, state.isDirty, 'Librevia', untitled)
    }

    sync()
    return useWorkspace.subscribe(sync)
  }, [])

  // A casca inteira muda de cor conforme o que está aberto — azul de
  // documento, verde de planilha. Ver o comentário de `--accent` no CSS.
  return (
    <div
      className={[
        'app',
        workbook === null ? '' : 'app--spreadsheet',
        // A casca inteira encolhe: e a classe que some com a barra de status
        // aqui embaixo e com a de ferramentas la dentro do editor.
        reading ? 'app--reading' : '',
      ]
        .filter((name) => name !== '')
        .join(' ')}
    >
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
      {/* A barra de status sai no modo de leitura: contagem de palavras e
          numero de paginas sao ferramentas de quem escreve. */}
      {hasFile && !reading && <StatusBar />}
    </div>
  )
}
