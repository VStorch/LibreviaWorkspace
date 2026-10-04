import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { JSONContent } from '@tiptap/core'
import type { Editor } from '@tiptap/react'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import type { ContextMenuTarget } from '@shared/types.js'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS, NOTES_CSS } from '@services/document/content-styles.js'
import { styleSheetCss } from '@services/document/style-css.js'
import { plainPasteContent } from '@services/document/paste.js'
import type { DocumentNode, PageSetup } from '@services/document/model.js'
import { effectiveSections, resolveSections, type ResolvedSections } from '@services/document/sections.js'
import { fitWidthZoom } from '@services/document/zoom.js'
import { usePreferences } from '../state/preferences.js'
import { useLeaveReadingOnEscape, useReadingMode } from '../state/reading.js'
import { useRevisionView } from '../state/revision-view.js'
import { t as translateNow, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { setFittedZoom, useEffectiveZoom } from '../state/zoom.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { NavigationPane } from './NavigationPane.js'
import { COMMENTS_PANE_WIDTH_PX } from './CommentsPane.js'
import { useComments } from './useComments.js'
import { usePagination, type PageLayout } from './usePagination.js'
import { useBandHeights } from './useBandHeights.js'
import { splitIntoPages } from './print-source.js'
import { marksOfDoc } from './section-commands.js'
import { useEditorCommands, type EditorCommands } from './useEditorCommands.js'
import { settlePageFields, type ReferenceContext } from './references.js'
import { footnotePagesOf, notePagesOf, samePages, setNotePages } from './extensions/note-ref.js'
import type { SearchStatus } from './extensions/search-replace.js'
import { useWorkspaceEditor } from './useWorkspaceEditor.js'
import { useEditorSync } from './useEditorSync.js'
import { useFloatsByPage, useSheetEditing } from './useSheetEditing.js'
import { useSheetGeometry } from './useSheetGeometry.js'
import { EditorContextMenu, EditorDialogs } from './EditorDialogs.js'
import { PageStack } from './PageStack.js'

/**
 * Pagina ao vivo: o texto é um fluxo só, e uma decoração de margem empurra cada
 * bloco para a folha certa (ver `extensions/pagination.ts`). As faixas são
 * desenhadas fora do `contenteditable`, na camada das folhas.
 */
export function DocumentEditor(): React.JSX.Element {
  const readOnly = useWorkspace((state) => state.readOnly)
  const styles = useWorkspace((state) => state.styles)
  const preferences = usePreferences((state) => state.preferences)
  const reading = useReadingMode()
  const revisionView = useRevisionView((state) => state.view)
  useLeaveReadingOnEscape(reading)

  const notePoolRef = useRef<HTMLDivElement>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const zoom = useEffectiveZoom()
  const { editor, searchStatus, contentRevision, touch } = useWorkspaceEditor()
  useEditorSync(editor, {
    readOnly,
    reading,
    revisionView,
    styles,
    preferences,
    notePool: notePoolRef,
    touch,
  })

  const { resolved, effective } = useDocumentSections(editor)
  const { layout, bands, layoutRef } = useDocumentLayout(editor, resolved, effective, contentRevision)
  const referenceContext = useReferenceContext(layoutRef, effective)
  const pasteWithoutFormat = usePasteWithoutFormat(editor, readOnly)
  const commands = useEditorCommands(editor, readOnly, pasteWithoutFormat, referenceContext)
  const { comments, outside } = useComments(editor)
  const geometry = useSheetGeometry({
    editor,
    effective,
    sections: resolved.sections,
    layout,
    bands,
    reading,
  })
  const floatsByPage = useFloatsByPage(editor, layout, contentRevision)
  const sheetEditing = useSheetEditing(editor, readOnly)
  useFittedZoom(scrollRef, editor, geometry.stackWidthPx)

  // O segundo passe dos campos de página (`settlePageFields`).
  useEffect(() => {
    if (editor !== null && !readOnly) settlePageFields(editor, referenceContext())
  }, [editor, layout, readOnly, referenceContext])

  if (editor === null) return <div className="editor-shell" />

  const commentsPane = !reading && preferences.commentsPane && comments.length > 0
  return (
    <div className="editor-shell">
      <EditorChrome
        editor={editor}
        commands={commands}
        searchStatus={searchStatus}
        resolved={resolved}
        readOnly={readOnly}
        referenceContext={referenceContext}
        pasteWithoutFormat={pasteWithoutFormat}
      />

      {reading && <ReadingHint />}

      {/* Ao lado da folha, e não por cima: um painel flutuante taparia o texto. */}
      <div className="editor-body">
        {!reading && preferences.navigationPane && <NavigationPane editor={editor} />}
        <div ref={scrollRef} className={`editor-scroll${reading ? ' editor-scroll--reading' : ''}`}>
          <PageStack
            editor={editor}
            layout={layout}
            reading={reading}
            zoom={zoom}
            geometry={geometry}
            effective={effective}
            floatsByPage={floatsByPage}
            sheetEditing={sheetEditing}
            notePool={notePoolRef}
            comments={commentsPane ? { list: comments, outside } : null}
            zoomedWidthPx={geometry.stackWidthPx + (commentsPane ? COMMENTS_PANE_WIDTH_PX : 0)}
          />
        </div>
      </div>
    </div>
  )
}

interface EditorChromeProps {
  readonly editor: Editor
  readonly commands: EditorCommands
  readonly searchStatus: SearchStatus
  readonly resolved: ResolvedSections
  readonly readOnly: boolean
  readonly referenceContext: () => ReferenceContext
  readonly pasteWithoutFormat: () => Promise<void>
}

/** O CSS de impressão, a barra, os diálogos abertos e o menu de contexto. */
function EditorChrome(props: EditorChromeProps): React.JSX.Element {
  const { editor, commands, pasteWithoutFormat } = props
  const { dialogs, setDialog } = commands
  const styles = useWorkspace((state) => state.styles)
  const styleCss = useMemo(() => styleSheetCss(styles), [styles])
  const showToolbar = usePreferences((state) => state.preferences.showToolbar)
  const reading = useReadingMode()
  const [contextTarget, setContextTarget] = useContextMenuTarget()
  return (
    <>
      {/* O mesmo CSS do HTML de impressão, para o papel sair igual à tela. */}
      <style>{DOCUMENT_CONTENT_CSS + styleCss + NOTES_CSS + EDITOR_ONLY_CSS}</style>

      {!reading && showToolbar && (
        <DocumentToolbar
          editor={editor}
          onOpenFind={() => setDialog('find', true)}
          onOpenPageSetup={() => setDialog('pageSetup', true)}
          onOpenStyles={() => setDialog('styles', true)}
          paragraphOpen={dialogs.paragraph}
          onParagraphOpenChange={(paragraph) => setDialog('paragraph', paragraph)}
          onOpenTable={() => setDialog('table', true)}
          onOpenImageProperties={() => setDialog('imageProperties', true)}
          onOpenListFormat={() => setDialog('listFormat', true)}
        />
      )}

      <EditorDialogs
        open={dialogs}
        context={{
          editor,
          close: (dialog) => () => setDialog(dialog, false),
          searchStatus: props.searchStatus,
          resolved: props.resolved,
          equationTarget: commands.equationTarget,
          readOnly: props.readOnly,
          referenceContext: props.referenceContext,
        }}
      />

      {contextTarget !== null && (
        <EditorContextMenu
          editor={editor}
          target={contextTarget}
          run={commands.run}
          setDialog={setDialog}
          pasteWithoutFormat={pasteWithoutFormat}
          onClose={() => setContextTarget(null)}
        />
      )}
    </>
  )
}

function ReadingHint(): React.JSX.Element {
  const t = useT()
  return (
    <div className="reading-hint" role="status">
      {t('view.reading.hint')}
    </div>
  )
}

/**
 * Pelas marcas, e não pelo documento: seções novas a cada tecla refariam a
 * paginação e a geometria das seções, até o React desistir (erro 185).
 */
function useDocumentSections(editor: Editor | null): {
  resolved: ResolvedSections
  effective: readonly PageSetup[]
} {
  const declaredPage = useWorkspace((state) => state.page)
  const library = useWorkspace((state) => state.sections)
  const doc = editor?.state.doc ?? null
  const marksKey = doc === null ? '[]' : JSON.stringify(marksOfDoc(doc))
  const bodySection: unknown = doc?.attrs['bodySection'] ?? null
  const resolved = useMemo(
    () => resolveSections(JSON.parse(marksKey) as (string | null)[], bodySection, declaredPage, library),
    [marksKey, bodySection, declaredPage, library],
  )
  const effective = useMemo(() => effectiveSections(resolved.page, resolved.sections), [resolved])
  return { resolved, effective }
}

/**
 * A paginação, as faixas e o que a impressão lê deles. Lidos por `ref` ao
 * imprimir: registrar `readPages` a cada layout recriaria a fonte do documento a
 * cada tecla.
 */
function useDocumentLayout(
  editor: Editor | null,
  resolved: ResolvedSections,
  effective: readonly PageSetup[],
  contentRevision: number,
): {
  layout: PageLayout
  bands: ReturnType<typeof useBandHeights>
  layoutRef: { readonly current: PageLayout }
} {
  const styles = useWorkspace((state) => state.styles)
  const reading = useReadingMode()
  const [sheetSections, setSheetSections] = useState('')
  const bands = useBandHeights(effective, contentRevision, sheetSections)
  const layout = usePagination(editor, effective, resolved.sections, contentRevision, {
    bands,
    paginated: !reading,
    styles,
  })
  const sheetSectionsNow = layout.sheets.map((sheet) => sheet.section).join(',')
  useEffect(() => {
    if (sheetSectionsNow !== sheetSections) setSheetSections(sheetSectionsNow)
  }, [sheetSectionsNow, sheetSections])

  const layoutRef = useRef(layout)
  layoutRef.current = layout
  const bandsRef = useRef(bands)
  bandsRef.current = bands

  const registerDocumentSource = useWorkspace((state) => state.registerDocumentSource)
  useEffect(() => {
    if (editor === null) return undefined
    registerDocumentSource({
      readDoc: () => editor.getJSON() as DocumentNode,
      readHtml: () => editor.getHTML(),
      readPages: () => ({
        pages: splitIntoPages(editor, layoutRef.current, effective),
        bands: bandsRef.current,
        sections: layoutRef.current.sheets.map((sheet) => sheet.section),
      }),
    })
    return () => registerDocumentSource(null)
  }, [editor, effective, registerDocumentSource])

  const setPageCount = useWorkspace((state) => state.setPageCount)
  useEffect(() => setPageCount(layout.pages), [layout.pages, setPageCount])
  useFootnotePages(editor, layout)
  return { layout, bands, layoutRef }
}

/** A folha de cada nota de rodapé, para o reinício por página. Transação sem mudança no texto. */
function useFootnotePages(editor: Editor | null, layout: PageLayout): void {
  const notesSetup = useWorkspace((state) => state.notes)
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    const current = notePagesOf(editor.state)
    const pages = notesSetup?.footnotePr?.restart === 'eachPage' ? footnotePagesOf(layout.noteAreas) : []
    if (samePages(pages, current)) return
    editor.view.dispatch(setNotePages(editor.state.tr, pages))
  }, [editor, layout.noteAreas, notesSetup])
}

/** Lidas na hora do comando, e não na da renderização. */
function useReferenceContext(
  layoutRef: { readonly current: PageLayout },
  effective: readonly PageSetup[],
): () => ReferenceContext {
  const effectiveRef = useRef(effective)
  effectiveRef.current = effective
  return useCallback(
    (): ReferenceContext => ({
      layout: layoutRef.current,
      page: useWorkspace.getState().page,
      sections: effectiveRef.current,
      styles: useWorkspace.getState().styles,
      setStyles: useWorkspace.getState().setStyles,
      outsideBookmarks: useWorkspace.getState().outsideBookmarks,
      t: translateNow,
    }),
    [layoutRef],
  )
}

/**
 * Nada de `pasteAndMatchStyle` do Chromium: ele **adapta** a formatação em vez
 * de descartá-la. A conversão é `@services/document/paste.ts`.
 */
function usePasteWithoutFormat(editor: Editor | null, readOnly: boolean): () => Promise<void> {
  const showError = useWorkspace((state) => state.showError)
  return useCallback(async (): Promise<void> => {
    if (editor === null || readOnly) return

    const result = await window.api.edit.readClipboardText({})
    if (!result.ok) {
      showError(result.error)
      return
    }

    const content = plainPasteContent(result.data.text)
    if (content.length === 0) return

    // `DocumentNode` e `JSONContent` são a mesma forma.
    editor
      .chain()
      .focus()
      .insertContent(content as JSONContent[])
      .run()
  }, [editor, readOnly, showError])
}

function useContextMenuTarget(): [ContextMenuTarget | null, (target: ContextMenuTarget | null) => void] {
  const [contextTarget, setContextTarget] = useState<ContextMenuTarget | null>(null)
  useEffect(
    () =>
      window.api.contextMenu.onRequest((payload) => {
        // A segunda ponta do zod, que o preload sandboxed não pode fazer.
        const parsed = pushContracts[IpcChannel.ContextMenuRequested].safeParse(payload)
        if (!parsed.success) return

        // Sem ação a oferecer, o menu teria todos os itens apagados, como na barra de ferramentas.
        if (!parsed.data.editable && !parsed.data.canCopy) return

        setContextTarget(parsed.data)
      }),
    [],
  )
  return [contextTarget, setContextTarget]
}

/** Medido sempre: ampliar a partir do ajuste precisa do valor que se vê. */
function useFittedZoom(
  scrollRef: { readonly current: HTMLDivElement | null },
  editor: Editor | null,
  stackWidthPx: number,
): void {
  useEffect(() => {
    const scroll = scrollRef.current
    if (scroll === null) return undefined
    const measure = (): void => setFittedZoom(fitWidthZoom(scroll.clientWidth, stackWidthPx))
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(scroll)
    return () => observer.disconnect()
  }, [scrollRef, editor, stackWidthPx])
}
