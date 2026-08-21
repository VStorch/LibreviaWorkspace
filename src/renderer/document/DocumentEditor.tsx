import { useCallback, useEffect, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS } from '@services/document/content-styles.js'
import { hasBandContent, mmToPx, pageDimensionsMm } from '@services/document/model.js'
import type { DocumentNode } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageBand } from './PageBand.js'
import { PageGuides, usePageBreaks } from './PageGuides.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { onEditorCommand } from './editor-commands.js'
import type { SearchStatus } from './extensions/search-replace.js'

/**
 * Editor de documentos.
 *
 * A moldura de página é visual: largura e margens em milímetros, rolagem
 * contínua. O editor não pagina ao vivo — essa decisão está registrada em
 * docs/00-plano-tecnico.md §6.3, e a paginação real acontece na exportação
 * para PDF, na Fase 3.
 */
export function DocumentEditor(): React.JSX.Element {
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const page = useWorkspace((state) => state.page)
  const markDirty = useWorkspace((state) => state.markDirty)
  const setStats = useWorkspace((state) => state.setStats)
  const registerDocumentSource = useWorkspace((state) => state.registerDocumentSource)
  const setEstimatedPages = useWorkspace((state) => state.setEstimatedPages)
  const readOnly = useWorkspace((state) => state.readOnly)

  const pageRef = useRef<HTMLDivElement>(null)
  const [contentRevision, setContentRevision] = useState(0)

  const [searchStatus, setSearchStatus] = useState<SearchStatus>({ total: 0, current: 0 })
  const [findOpen, setFindOpen] = useState(false)
  const [pageSetupOpen, setPageSetupOpen] = useState(false)

  const handleSearchStatus = useCallback((status: SearchStatus) => setSearchStatus(status), [])

  const editor = useEditor({
    extensions: buildEditorExtensions(handleSearchStatus),
    content: initialDoc,
    onUpdate: ({ editor: current }) => {
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
      attributes: { class: 'page__content', spellcheck: 'false' },
    },
  })

  /**
   * Somente leitura ligado e desligado no editor já montado.
   *
   * Passar `editable` na criação não bastaria: liberar a edição pelo aviso
   * acontece **depois**, e recriar o editor ali perderia a posição do cursor e
   * o histórico de desfazer.
   *
   * O segundo argumento é o que importa: `setEditable` emite um update por
   * padrão, e o update marca o documento como alterado. Sem ele, todo arquivo
   * aberto aparecia como "não salvo" antes de o usuário tocar em nada — e o
   * aviso de descarte apareceria ao fechar um documento que ninguém editou.
   */
  useEffect(() => {
    editor?.setEditable(!readOnly, false)
  }, [editor, readOnly])

  // Salvar e imprimir precisam do conteúdo atual, que só o editor conhece.
  useEffect(() => {
    if (editor === null) return undefined
    registerDocumentSource({
      readDoc: () => editor.getJSON() as DocumentNode,
      readHtml: () => editor.getHTML(),
    })
    return () => registerDocumentSource(null)
  }, [editor, registerDocumentSource])

  useEffect(
    () =>
      onEditorCommand((command) => {
        if (command === 'find-replace') setFindOpen(true)
        if (command === 'page-setup') setPageSetupOpen(true)
        if (command === 'insert-page-break') editor?.chain().focus().setPageBreak().run()
      }),
    [editor],
  )

  const pageBreaks = usePageBreaks(pageRef, page, contentRevision)

  useEffect(() => setEstimatedPages(pageBreaks.length + 1), [pageBreaks, setEstimatedPages])

  if (editor === null) return <div className="editor-shell" />

  const { width, height } = pageDimensionsMm(page)
  // Metade da margem: a faixa do cabeçalho é mais larga que a coluna de texto,
  // como no documento original.
  const bandInset = mmToPx(Math.min(page.margins.left, page.margins.right) / 2)

  return (
    <div className="editor-shell">
      {/* O estilo do conteúdo vem do mesmo módulo que o HTML de impressão usa.
          Duas folhas de estilo divergiriam com o tempo, e o PDF deixaria de
          sair igual à tela — o risco registrado no §6.3 do plano. */}
      <style>{DOCUMENT_CONTENT_CSS + EDITOR_ONLY_CSS}</style>

      <DocumentToolbar
        editor={editor}
        onOpenFind={() => setFindOpen(true)}
        onOpenPageSetup={() => setPageSetupOpen(true)}
      />

      {findOpen && (
        <FindReplacePanel editor={editor} status={searchStatus} onClose={() => setFindOpen(false)} />
      )}

      {pageSetupOpen && <PageSetupPanel onClose={() => setPageSetupOpen(false)} />}

      <div className="editor-scroll">
        <div
          ref={pageRef}
          className="page"
          style={{
            width: `${mmToPx(width)}px`,
            // Uma folha vazia tem a altura de uma folha, não a do texto que já
            // existe nela. Sem isto o documento novo abre como uma tira baixa,
            // e a noção de papel — que é o que o Writer, o Word e o Docs dão —
            // se perde. Continua sendo moldura: quem pagina é a exportação.
            minHeight: `${mmToPx(height)}px`,
            paddingTop: `${mmToPx(page.margins.top)}px`,
            paddingRight: `${mmToPx(page.margins.right)}px`,
            paddingBottom: `${mmToPx(page.margins.bottom)}px`,
            paddingLeft: `${mmToPx(page.margins.left)}px`,
          }}
        >
          <PageGuides offsets={pageBreaks} topOffsetPx={mmToPx(page.margins.top)} />

          {/* As faixas moram dentro da margem, como no papel: por isso são
              posicionadas em relação à folha e não empurram o texto. */}
          {hasBandContent(page.headerBand) && (
            <PageBand band={page.headerBand} kind="header" pageNumber={1} insetPx={bandInset} />
          )}

          <EditorContent editor={editor} />

          {hasBandContent(page.footerBand) && (
            <PageBand band={page.footerBand} kind="footer" pageNumber={1} insetPx={bandInset} />
          )}
        </div>
      </div>
    </div>
  )
}
