import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { JSONContent } from '@tiptap/core'
import { EditorContent, useEditor } from '@tiptap/react'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import { RevisionView, type ContextMenuTarget } from '@shared/types.js'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS, NOTES_CSS } from '@services/document/content-styles.js'
import { styleSheetCss } from '@services/document/style-css.js'
import { plainPasteContent } from '@services/document/paste.js'
import {
  contentInsetsMm,
  mmToPx,
  pageDimensionsMm,
  pxToMm,
  type DocumentNode,
  type PageSetup,
} from '@services/document/model.js'
import { editBandFloat, editBandPiece } from '@services/document/band.js'
import {
  columnGeometry,
  effectiveSections,
  resolveSections,
  sheetSetups,
} from '@services/document/sections.js'
import { NO_BANDS } from '@services/document/band.js'
import { floatsOf } from '@services/document/floating.js'
import { currentPreferences, usePreferences } from '../state/preferences.js'
import { useLeaveReadingOnEscape, useReadingMode } from '../state/reading.js'
import { useRevisionView } from '../state/revision-view.js'
import { revisionViewOf, setRevisionViewMeta } from './extensions/revision-view.js'
import { t as translateNow, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { setFittedZoom, useEffectiveZoom } from '../state/zoom.js'
import { fitWidthZoom } from '@services/document/zoom.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { TableDialog } from './toolbar/TableDialog.js'
import { TablePropertiesDialog } from './toolbar/TablePropertiesDialog.js'
import { ImageDialog } from './toolbar/ImageDialog.js'
import { DocumentContextMenu } from './DocumentContextMenu.js'
import { ListFormatDialog, ListStartDialog } from './ListFormatDialog.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { SpecialCharsDialog } from './SpecialCharsDialog.js'
import { MathDialog } from './MathDialog.js'
import { StylesPanel } from './StylesPanel.js'
import { NavigationPane } from './NavigationPane.js'
import { BookmarkDialog } from './BookmarkDialog.js'
import { ColumnsDialog } from './ColumnsDialog.js'
import { CaptionDialog } from './CaptionDialog.js'
import { CrossReferenceDialog } from './CrossReferenceDialog.js'
import { WordCountDialog } from './WordCountDialog.js'
import { PropertiesDialog } from './PropertiesDialog.js'
import { AuthorNameDialog } from './AuthorNameDialog.js'
import { PaperSheet } from './PaperSheet.js'
import { noteBodiesOf, setNotePool } from './extensions/note-view.js'
import { COMMENTS_PANE_WIDTH_PX, CommentsPane } from './CommentsPane.js'
import { useComments } from './useComments.js'
import { insertComment } from './comment-commands.js'
import { hasChangeAtCursor, settleChange } from './revision-commands.js'
import { convertNote, noteAtCursor } from './note-commands.js'
import { focusComment } from './extensions/comment.js'
import { usePagination } from './usePagination.js'
import { useBandHeights } from './useBandHeights.js'
import { splitIntoPages } from './print-source.js'
import type { FloatSource, PlacedFloat } from './FloatingLayer.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { isPaginationOnly } from './extensions/pagination.js'
import { setSectionBoxes } from './extensions/section-geometry.js'
import { marksOfDoc, sectionAtCursor } from './section-commands.js'
import { useEditorCommands } from './useEditorCommands.js'
import { settlePageFields, type ReferenceContext } from './references.js'
import { footnotePagesOf, notePagesOf, samePages, setNotePages } from './extensions/note-ref.js'
import type { SearchStatus } from './extensions/search-replace.js'

/**
 * Pagina ao vivo: o texto é um fluxo só, e uma decoração de margem empurra cada
 * bloco para a folha certa (ver `extensions/pagination.ts`). As faixas são
 * desenhadas fora do `contenteditable`, na camada das folhas.
 */
export function DocumentEditor(): React.JSX.Element {
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const declaredPage = useWorkspace((state) => state.page)
  const library = useWorkspace((state) => state.sections)
  const markDirty = useWorkspace((state) => state.markDirty)
  const setStats = useWorkspace((state) => state.setStats)
  const registerDocumentSource = useWorkspace((state) => state.registerDocumentSource)
  const setPageCount = useWorkspace((state) => state.setPageCount)
  const readOnly = useWorkspace((state) => state.readOnly)
  const setPage = useWorkspace((state) => state.setPage)
  const setSections = useWorkspace((state) => state.setSections)
  const styles = useWorkspace((state) => state.styles)
  const styleCss = useMemo(() => styleSheetCss(styles), [styles])
  const showError = useWorkspace((state) => state.showError)
  const preferences = usePreferences((state) => state.preferences)
  const reading = useReadingMode()
  const revisionView = useRevisionView((state) => state.view)
  const t = useT()

  useLeaveReadingOnEscape(reading)

  const pageRef = useRef<HTMLDivElement>(null)
  const notePoolRef = useRef<HTMLDivElement>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const zoom = useEffectiveZoom()
  const [contentRevision, setContentRevision] = useState(0)

  const [searchStatus, setSearchStatus] = useState<SearchStatus>({ total: 0, current: 0 })

  const [contextTarget, setContextTarget] = useState<ContextMenuTarget | null>(null)

  const handleSearchStatus = useCallback((status: SearchStatus) => setSearchStatus(status), [])

  const editor = useEditor({
    // Da loja, e não das props: o editor é criado uma vez, e o resto chega pelos efeitos.
    extensions: buildEditorExtensions(handleSearchStatus, {
      isTypographyEnabled: () => currentPreferences().typography,
      invisibleCharactersVisible: currentPreferences().invisibleCharacters,
      isKnownComment: (cid) => useWorkspace.getState().comments.some((comment) => comment.id === cid),
      notes: () => useWorkspace.getState().notes,
      isTrackingChanges: () => useWorkspace.getState().trackChanges === true,
      // O `w:author` é obrigatório: sem nome, um autor genérico, como no Word.
      revisionAuthor: () => currentPreferences().authorName.trim() || translateNow('revisions.unknownAuthor'),
    }),
    content: initialDoc,
    onUpdate: ({ editor: current, transaction }) => {
      // A paginação chega como transação: tratá-la como edição sujaria o documento.
      if (isPaginationOnly(transaction)) return

      markDirty()
      setContentRevision((value) => value + 1)
      setStats({
        characters: current.storage['characterCount'].characters(),
        words: current.storage['characterCount'].words(),
      })
    },
    onCreate: ({ editor: current }) => {
      setStats({
        characters: current.storage['characterCount'].characters(),
        words: current.storage['characterCount'].words(),
      })
    },
    editorProps: {
      attributes: {
        class: 'page__content',
        spellcheck: currentPreferences().spellcheck ? 'true' : 'false',
      },
    },
  })

  // Pelas marcas, e não pelo documento: seções novas a cada tecla refariam a
  // paginação e a geometria das seções, até o React desistir (erro 185).
  const doc = editor?.state.doc ?? null
  const marksKey = doc === null ? '[]' : JSON.stringify(marksOfDoc(doc))
  const bodySection: unknown = doc?.attrs['bodySection'] ?? null
  const resolved = useMemo(
    () => resolveSections(JSON.parse(marksKey) as (string | null)[], bodySection, declaredPage, library),
    [marksKey, bodySection, declaredPage, library],
  )
  const sections = resolved.sections
  const effective = useMemo(() => effectiveSections(resolved.page, sections), [resolved.page, sections])
  const page = effective.at(-1)!
  // Lidas na hora do comando, e não na da renderização.
  const effectiveRef = useRef(effective)
  effectiveRef.current = effective

  // `setEditable` no editor já montado, e com o segundo argumento: recriar o
  // editor perderia cursor e histórico, e o update padrão marcaria todo arquivo
  // aberto como "não salvo". O modo de leitura e o Original também travam a edição.
  const original = revisionView === RevisionView.Original
  useEffect(() => {
    editor?.setEditable(!readOnly && !reading && !original, false)
    if (editor !== null && !editor.isDestroyed) for (const body of noteBodiesOf(editor.view)) body.refresh()
  }, [editor, readOnly, reading, original])

  useEffect(() => {
    if (editor === null || editor.isDestroyed) return undefined
    const view = editor.view
    setNotePool(view, notePoolRef.current)
    return () => setNotePool(view, null)
  }, [editor])

  // Uma transação sem mudança no documento: a paginação mede de novo, e o escondido não ocupa lugar.
  useEffect(() => {
    if (editor === null || editor.isDestroyed || revisionViewOf(editor.state) === revisionView) return
    editor.view.dispatch(setRevisionViewMeta(editor.state.tr, revisionView))
    // O CSS esconde pela classe do modo (`revisions-…`), que o corpo das notas também leva.
    for (const body of noteBodiesOf(editor.view)) body.refresh()
    setContentRevision((value) => value + 1)
  }, [editor, revisionView])

  useEffect(() => {
    if (editor !== null) editor.storage.paragraphCommands.styles = styles
  }, [editor, styles])

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

  /**
   * Escrito no elemento: o ProseMirror só lê os atributos ao criar a visão. O
   * main liga o corretor na sessão; isto é a outra metade.
   */
  useEffect(() => {
    editor?.view.dom.setAttribute('spellcheck', preferences.spellcheck ? 'true' : 'false')
  }, [editor, preferences.spellcheck])

  // A transação das marcas de formatação não muda o documento.
  useEffect(() => {
    editor?.commands.showInvisibleCharacters(preferences.invisibleCharacters)
  }, [editor, preferences.invisibleCharacters])

  // Sem o painel, sai só o realce: pontas e corpos ficam, e voltam ao arquivo.
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    editor.view.dispatch(focusComment(editor.state.tr, { hidden: !preferences.commentsPane }))
  }, [editor, preferences.commentsPane])

  /**
   * Nada de `pasteAndMatchStyle` do Chromium: ele **adapta** a formatação em vez
   * de descartá-la. A conversão é `@services/document/paste.ts`.
   */
  const pasteWithoutFormat = useCallback(async (): Promise<void> => {
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

  // A trava do somente leitura é conferida no `run` (`useEditorCommands`).
  const referenceContext = useCallback(
    (): ReferenceContext => ({
      layout: layoutRef.current,
      page: useWorkspace.getState().page,
      sections: effectiveRef.current,
      styles: useWorkspace.getState().styles,
      setStyles: useWorkspace.getState().setStyles,
      outsideBookmarks: useWorkspace.getState().outsideBookmarks,
      t: translateNow,
    }),
    [],
  )

  const { dialogs, setDialog, run, equationTarget } = useEditorCommands(
    editor,
    readOnly,
    pasteWithoutFormat,
    referenceContext,
  )

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

  const { comments, outside } = useComments(editor)
  const [sheetSections, setSheetSections] = useState('')
  const bands = useBandHeights(effective, contentRevision, sheetSections)
  const layout = usePagination(editor, effective, sections, contentRevision, {
    bands,
    paginated: !reading,
    styles,
  })
  const sheetSectionsNow = layout.sheets.map((sheet) => sheet.section).join(',')
  useEffect(() => {
    if (sheetSectionsNow !== sheetSections) setSheetSections(sheetSectionsNow)
  }, [sheetSectionsNow, sheetSections])
  const sheetSetupList = useMemo(() => sheetSetups(effective, layout.sheets), [effective, layout.sheets])

  // Cada folha centrada na pilha, como o Word mostra retrato e paisagem juntos.
  const stackWidthPx = Math.max(layout.stackWidthPx, mmToPx(pageDimensionsMm(page).width))
  // O invólucro do zoom cresce para a rolagem chegar até a coluna dos comentários.
  const commentsPane = !reading && preferences.commentsPane && comments.length > 0
  const zoomedWidthPx = stackWidthPx + (commentsPane ? COMMENTS_PANE_WIDTH_PX : 0)
  const firstSection = effective[layout.sheets[0]?.section ?? 0] ?? page
  const insets = contentInsetsMm(firstSection, bands[layout.sheets[0]?.section ?? 0] ?? NO_BANDS)
  const baseLeftPx = (stackWidthPx - mmToPx(pageDimensionsMm(page).width)) / 2 + mmToPx(page.margins.left)
  const baseRightPx = (stackWidthPx - mmToPx(pageDimensionsMm(page).width)) / 2 + mmToPx(page.margins.right)

  // No modo de leitura não há folha, e nada se desloca.
  useEffect(() => {
    if (editor === null) return
    const boxes = reading
      ? []
      : effective.map((section) => {
          const widthPx = mmToPx(pageDimensionsMm(section).width)
          const left = (stackWidthPx - widthPx) / 2 + mmToPx(section.margins.left)
          // Qual coluna, quem decide é a paginação, por translação.
          const content = mmToPx(columnGeometry(section).widthMm)
          const base = stackWidthPx - baseLeftPx - baseRightPx
          return { shiftPx: left - baseLeftPx, narrowerPx: base - content }
        })
    setSectionBoxes(editor.view, boxes, sections)
  }, [editor, effective, sections, reading, stackWidthPx, baseLeftPx, baseRightPx])

  // A posição depende da folha em que o parágrafo âncora caiu.
  const floatsByPage = useMemo(() => {
    const pages: PlacedFloat[][] = Array.from({ length: layout.pages }, () => [])
    if (editor === null) return pages

    let index = 0
    editor.state.doc.forEach((node, pos) => {
      const anchor = layout.anchors[index]
      index += 1
      if (anchor === undefined) return

      const sheet = pages[anchor.pageIndex]
      if (sheet === undefined) return

      let slot = 0
      for (const object of floatsOf(node.attrs)) {
        // É pela posição do bloco que o texto da caixa volta ao atributo de onde saiu.
        sheet.push({ object, anchorTopMm: pxToMm(anchor.topPx), source: { pos, index: slot } })
        slot += 1
      }
    })

    return pages
  }, [editor, layout, contentRevision])

  /** Uma transação comum: entra no histórico, suja o documento e chega ao gravador pelo `getJSON()`. */
  const editFloat = useCallback(
    (source: FloatSource, content: DocumentNode[]) => {
      if (editor === null || readOnly) return

      const node = editor.state.doc.nodeAt(source.pos)
      if (node === null) return

      const floats = node.attrs['floats']
      if (!Array.isArray(floats)) return

      const object = floats[source.index] as { content?: unknown } | undefined
      if (object === undefined) return
      if (JSON.stringify(object.content ?? []) === JSON.stringify(content)) return

      const updated = floats.map((item, index) => (index === source.index ? { ...object, content } : item))
      editor.view.dispatch(editor.state.tr.setNodeAttribute(source.pos, 'floats', updated))
    },
    [editor, readOnly],
  )

  /** A faixa mora em `page`, e não no documento do editor: por isso `setPage`, e não transação. */
  const editAllSections = useCallback(
    (change: <T extends PageSetup>(section: T) => T) => {
      // Da loja: várias peças podem sair do foco em sequência.
      const state = useWorkspace.getState()
      const page = change(state.page)
      if (page !== state.page) setPage(page)
      const sections = state.sections.map(change)
      if (sections.some((section, index) => section !== state.sections[index])) setSections(sections)
    },
    [setPage, setSections],
  )

  const editBand = useCallback(
    (pid: string, text: string) => {
      if (readOnly) return

      // A quebra de seção copia as referências, e a mesma parte mora nas duas seções.
      editAllSections((section) => editBandPiece(section, pid, text))
    },
    [readOnly, editAllSections],
  )

  /** A caixa vem inteira: digitar dentro dela abre e fecha parágrafos. */
  const editBandBox = useCallback(
    (bid: string, content: DocumentNode[]) => {
      if (readOnly) return

      editAllSections((section) => editBandFloat(section, bid, content))
    },
    [readOnly, editAllSections],
  )

  // Lido ao imprimir: registrar `readPages` a cada layout recriaria a fonte do documento a cada tecla.
  const layoutRef = useRef(layout)
  layoutRef.current = layout

  const bandsRef = useRef(bands)
  bandsRef.current = bands

  useEffect(() => setPageCount(layout.pages), [layout.pages, setPageCount])

  // A folha de cada nota de rodapé, para o reinício por página. Transação sem mudança no texto.
  const notesSetup = useWorkspace((state) => state.notes)
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    const current = notePagesOf(editor.state)
    const pages = notesSetup?.footnotePr?.restart === 'eachPage' ? footnotePagesOf(layout.noteAreas) : []
    if (samePages(pages, current)) return
    editor.view.dispatch(setNotePages(editor.state.tr, pages))
  }, [editor, layout.noteAreas, notesSetup])

  // O segundo passe dos campos de página (`settlePageFields`).
  useEffect(() => {
    if (editor !== null && !readOnly) settlePageFields(editor, referenceContext())
  }, [editor, layout, readOnly, referenceContext])

  // Medido sempre: ampliar a partir do ajuste precisa do valor que se vê.
  useEffect(() => {
    const scroll = scrollRef.current
    if (scroll === null) return undefined
    const measure = (): void => setFittedZoom(fitWidthZoom(scroll.clientWidth, stackWidthPx))
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(scroll)
    return () => observer.disconnect()
  }, [editor, stackWidthPx])

  if (editor === null) return <div className="editor-shell" />

  const sheetBox = (index: number): { leftPx: number; widthPx: number; heightPx: number } => {
    const setup = effective[layout.sheets[index]?.section ?? effective.length - 1] ?? page
    const { width, height } = pageDimensionsMm(setup)
    return { leftPx: (stackWidthPx - mmToPx(width)) / 2, widthPx: mmToPx(width), heightPx: mmToPx(height) }
  }

  const editableSheet = readOnly
    ? {}
    : { onEditFloat: editFloat, onEditBandPiece: editBand, onEditBandBox: editBandBox }

  const contextNote = contextTarget === null ? null : noteAtCursor(editor)

  return (
    <div className="editor-shell">
      {/* O mesmo CSS do HTML de impressão, para o papel sair igual à tela. */}
      <style>{DOCUMENT_CONTENT_CSS + styleCss + NOTES_CSS + EDITOR_ONLY_CSS}</style>

      {!reading && preferences.showToolbar && (
        <DocumentToolbar
          editor={editor}
          onOpenFind={() => setDialog('find', true)}
          onOpenPageSetup={() => setDialog('pageSetup', true)}
          onOpenStyles={() => setDialog('styles', true)}
          paragraphOpen={dialogs.paragraph}
          onParagraphOpenChange={(open) => setDialog('paragraph', open)}
          onOpenTable={() => setDialog('table', true)}
          onOpenImageProperties={() => setDialog('imageProperties', true)}
          onOpenListFormat={() => setDialog('listFormat', true)}
        />
      )}

      {dialogs.find && (
        <FindReplacePanel editor={editor} status={searchStatus} onClose={() => setDialog('find', false)} />
      )}

      {dialogs.pageSetup && (
        <PageSetupPanel
          onClose={() => setDialog('pageSetup', false)}
          resolved={resolved}
          sectionIndex={sectionAtCursor(editor, resolved)}
        />
      )}

      {dialogs.wordCount && <WordCountDialog editor={editor} onClose={() => setDialog('wordCount', false)} />}

      {dialogs.properties && (
        <PropertiesDialog editor={editor} onClose={() => setDialog('properties', false)} />
      )}

      {dialogs.authorName && (
        <AuthorNameDialog editor={editor} onClose={() => setDialog('authorName', false)} />
      )}

      {dialogs.styles && <StylesPanel editor={editor} onClose={() => setDialog('styles', false)} />}

      {dialogs.specialCharacter && (
        <SpecialCharsDialog editor={editor} onClose={() => setDialog('specialCharacter', false)} />
      )}

      {dialogs.equation && (
        <MathDialog
          editor={editor}
          target={equationTarget}
          readOnly={readOnly}
          onClose={() => setDialog('equation', false)}
        />
      )}

      {dialogs.table && <TableDialog editor={editor} onClose={() => setDialog('table', false)} />}

      {dialogs.tableProperties && (
        <TablePropertiesDialog editor={editor} onClose={() => setDialog('tableProperties', false)} />
      )}

      {dialogs.imageProperties && (
        <ImageDialog editor={editor} onClose={() => setDialog('imageProperties', false)} />
      )}

      {dialogs.listFormat && (
        <ListFormatDialog editor={editor} onClose={() => setDialog('listFormat', false)} />
      )}

      {dialogs.listStart && <ListStartDialog editor={editor} onClose={() => setDialog('listStart', false)} />}

      {dialogs.columns && <ColumnsDialog editor={editor} onClose={() => setDialog('columns', false)} />}

      {dialogs.bookmark && <BookmarkDialog editor={editor} onClose={() => setDialog('bookmark', false)} />}

      {dialogs.caption && (
        <CaptionDialog
          editor={editor}
          context={referenceContext}
          onClose={() => setDialog('caption', false)}
        />
      )}

      {dialogs.crossReference && (
        <CrossReferenceDialog
          editor={editor}
          context={referenceContext}
          onClose={() => setDialog('crossReference', false)}
        />
      )}

      {contextTarget !== null && (
        <DocumentContextMenu
          target={contextTarget}
          // Fora de uma tabela, "mesclar células" não tem o que mesclar.
          inTable={editor.isActive('table')}
          onTableAction={run}
          // Como no Word, pelo botão direito sobre o item que se quer reiniciar.
          inList={
            editor.isActive('orderedList')
              ? 'orderedList'
              : editor.isActive('bulletList')
                ? 'bulletList'
                : null
          }
          onListAction={(action) => {
            if (action === 'restart') editor.chain().focus().restartListNumbering(1).run()
            else if (action === 'continue') editor.chain().focus().continueListNumbering().run()
            else setDialog(action === 'setStart' ? 'listStart' : 'listFormat', true)
          }}
          onClose={() => setContextTarget(null)}
          onPasteWithoutFormat={() => void pasteWithoutFormat()}
          onNewComment={() => insertComment(editor)}
          onRevision={hasChangeAtCursor(editor) ? (accept) => void settleChange(editor, accept) : null}
          noteKind={contextNote?.kind ?? null}
          onConvertNote={() => {
            if (contextNote !== null) convertNote(editor, contextNote.pos)
          }}
        />
      )}

      {reading && (
        <div className="reading-hint" role="status">
          {t('view.reading.hint')}
        </div>
      )}

      {/* Ao lado da folha, e não por cima: um painel flutuante taparia o texto. */}
      <div className="editor-body">
        {!reading && preferences.navigationPane && <NavigationPane editor={editor} />}
        <div ref={scrollRef} className={`editor-scroll${reading ? ' editor-scroll--reading' : ''}`}>
          {/* O `transform` do zoom não muda o espaço ocupado: este invólucro o
            reserva, para a rolagem chegar ao fim. A paginação mede em 100 %. */}
          <div
            className={`pages-zoom${reading ? ' pages-zoom--reading' : ''}`}
            style={
              reading
                ? undefined
                : {
                    width: `${(zoomedWidthPx * zoom) / 100}px`,
                    height: `${(layout.stackHeightPx * zoom) / 100}px`,
                  }
            }
          >
            <div
              ref={pageRef}
              className={`pages${reading ? ' pages--reading' : ''}`}
              data-zoom={reading ? 100 : zoom}
              style={
                reading
                  ? undefined
                  : {
                      width: `${stackWidthPx}px`,
                      height: `${layout.stackHeightPx}px`,
                      ...(zoom === 100
                        ? {}
                        : { transform: `scale(${zoom / 100})`, transformOrigin: 'top left' }),
                    }
              }
            >
              {/* Fora do `contenteditable`: dentro, cada folha seria um nó
              selecionável. Sem folhas no modo de leitura, nem objetos ancorados,
              cuja posição é relativa a uma folha. */}
              {!reading &&
                layout.sheetTops.map((top, index) => {
                  const box = sheetBox(index)
                  return (
                    <div
                      key={top}
                      className={`paper${(layout.sheetHeights[index] ?? 0) > box.heightPx + 1 ? ' paper--oversized' : ''}${layout.sheets[index]?.blank === true ? ' paper--blank' : ''}`}
                      style={{
                        top: `${top}px`,
                        height: `${layout.sheetHeights[index] ?? box.heightPx}px`,
                        left: `${box.leftPx}px`,
                        width: `${box.widthPx}px`,
                        right: 'auto',
                      }}
                      data-section={layout.sheets[index]?.section ?? 0}
                      aria-hidden="true"
                    >
                      <span className="paper__number">{index + 1}</span>
                    </div>
                  )
                })}

              {/* No papel as faixas moram dentro da margem e não empurram o texto. */}
              {!reading &&
                layout.sheetTops.map((top, index) => {
                  const box = sheetBox(index)
                  const setup = sheetSetupList[index] ?? { page, inSection: index + 1 }
                  return (
                    <PaperSheet
                      key={`banda-${top}`}
                      page={setup.page}
                      pageNumber={setup.inSection}
                      totalPages={layout.pages}
                      topPx={top}
                      leftPx={box.leftPx}
                      section={layout.sheets[index]?.section ?? 0}
                      floats={floatsByPage[index] ?? []}
                      columnLines={layout.columnLines.filter((line) => line.sheet === index)}
                      noteAreas={layout.noteAreas.filter((area) => area.sheet === index)}
                      schema={editor.schema}
                      {...editableSheet}
                    />
                  )
                })}

              <div
                className="pages__column"
                style={
                  reading
                    ? // A largura da leitura vem do CSS, e não das margens do documento.
                      undefined
                    : {
                        // A margem de cima é um piso, como no Word.
                        paddingTop: `${mmToPx(insets.top)}px`,
                        paddingRight: `${baseRightPx}px`,
                        paddingLeft: `${baseLeftPx}px`,
                      }
                }
              >
                <EditorContent editor={editor} />
              </div>

              {/* Os corpos de nota sem folha, na largura da coluna, onde a paginação
                  os mede; no modo de leitura, aparecem aqui, depois do texto. */}
              <div
                ref={notePoolRef}
                className={`note-pool${reading ? ' note-pool--reading' : ''}`}
                style={
                  reading
                    ? undefined
                    : {
                        left: `${baseLeftPx}px`,
                        width: `${mmToPx(pageDimensionsMm(page).width - page.margins.left - page.margins.right)}px`,
                      }
                }
              />

              {commentsPane && (
                <CommentsPane editor={editor} comments={comments} outside={outside} leftPx={stackWidthPx} />
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
