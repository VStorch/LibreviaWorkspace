import { useCallback, useEffect, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS } from '@services/document/content-styles.js'
import { hasBandContent, mmToPx, pageDimensionsMm } from '@services/document/model.js'
import type { DocumentNode } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageBand } from './PageBand.js'
import { usePagination } from './usePagination.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { isPaginationOnly } from './extensions/pagination.js'
import { onEditorCommand } from './editor-commands.js'
import type { SearchStatus } from './extensions/search-replace.js'

/**
 * Editor de documentos.
 *
 * Pagina ao vivo: o texto é um fluxo só, e as folhas são desenhadas atrás dele
 * nas posições que a medição produz. Quem empurra cada bloco para a folha certa
 * é uma decoração de margem — ver `extensions/pagination.ts`, que explica por
 * que não é um espaçador de verdade.
 *
 * Cabeçalho e rodapé se repetem em cada folha, com o número da página. São
 * desenhados fora do `contenteditable`, na mesma camada das folhas: no papel
 * eles moram dentro da margem, e ali não empurram o texto nem entram na
 * seleção.
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
    onUpdate: ({ editor: current, transaction }) => {
      // A paginação também chega como transação. Tratá-la como edição sujaria o
      // documento sem ninguém digitar, e a medição que ela dispara pediria
      // outra medição — o laço fecharia aqui.
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

  const layout = usePagination(editor, page, contentRevision)

  useEffect(() => setEstimatedPages(layout.pages), [layout.pages, setEstimatedPages])

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
          className="pages"
          style={{ width: `${mmToPx(width)}px`, height: `${layout.stackHeightPx}px` }}
        >
          {/* As folhas: papel desenhado atrás do texto. Ficam fora do
              `contenteditable` de propósito — dentro dele, cada folha seria um
              nó que a pessoa conseguiria selecionar e apagar. */}
          {layout.sheetTops.map((top, index) => (
            <div
              key={top}
              className="paper"
              style={{ top: `${top}px`, height: `${mmToPx(height)}px` }}
              aria-hidden="true"
            >
              <span className="paper__number">{index + 1}</span>
            </div>
          ))}

          {/* Uma faixa por folha, com o número real. No papel elas moram dentro
              da margem, e é por isso que não empurram o texto. */}
          {layout.sheetTops.map((top, index) => (
            <div
              key={`banda-${top}`}
              className="paper-bands"
              style={{ top: `${top}px`, height: `${mmToPx(height)}px` }}
            >
              {hasBandContent(page.headerBand) && (
                <PageBand
                  band={page.headerBand}
                  kind="header"
                  pageNumber={index + 1}
                  totalPages={layout.pages}
                  insetPx={bandInset}
                />
              )}
              {hasBandContent(page.footerBand) && (
                <PageBand
                  band={page.footerBand}
                  kind="footer"
                  pageNumber={index + 1}
                  totalPages={layout.pages}
                  insetPx={bandInset}
                />
              )}
            </div>
          ))}

          <div
            className="pages__column"
            style={{
              paddingTop: `${mmToPx(page.margins.top)}px`,
              paddingRight: `${mmToPx(page.margins.right)}px`,
              paddingLeft: `${mmToPx(page.margins.left)}px`,
            }}
          >
            <EditorContent editor={editor} />
          </div>
        </div>
      </div>
    </div>
  )
}
