import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS } from '@services/document/content-styles.js'
import {
  bandForPage,
  contentInsetsMm,
  hasBandContent,
  mmToPx,
  pageDimensionsMm,
  pxToMm as toMm,
  type BandHeights,
  type PageSetup,
} from '@services/document/model.js'
import type { DocumentNode } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageBand } from './PageBand.js'
import { usePagination, type PageLayout } from './usePagination.js'
import { FloatingLayer, type PlacedFloat } from './FloatingLayer.js'
import { bandFloatsOf, floatsOf, type FloatingObject } from '@services/document/floating.js'
import { pxToMm } from '@services/document/model.js'
import type { PrintFloat, PrintPage } from '@services/document/print-pages.js'
import { DOMSerializer, Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model'
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
      readPages: () => ({ pages: splitIntoPages(editor, layoutRef.current), bands: bandsRef.current }),
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

  const bands = useBandHeights(page, contentRevision)
  const layout = usePagination(editor, page, contentRevision, bands)
  const insets = contentInsetsMm(page, bands)

  // Os objetos ancorados de cada folha. Recalculados junto com a paginação
  // porque a posição de um deles depende de em que folha o parágrafo âncora
  // caiu — e isso muda a cada linha digitada.
  const floatsByPage = useMemo(() => {
    const pages: PlacedFloat[][] = Array.from({ length: layout.pages }, () => [])
    if (editor === null) return pages

    let index = 0
    editor.state.doc.forEach((node) => {
      const anchor = layout.anchors[index]
      index += 1
      if (anchor === undefined) return

      const sheet = pages[anchor.pageIndex]
      if (sheet === undefined) return

      for (const object of floatsOf(node.attrs)) {
        sheet.push({ object, anchorTopMm: pxToMm(anchor.topPx) })
      }
    })

    return pages
  }, [editor, layout, contentRevision])

  // O recorte em páginas é lido no momento de imprimir, e não no da renderização
  // — daí a `ref`: registrar `readPages` a cada mudança de layout recriaria a
  // fonte do documento dezenas de vezes por segundo enquanto se digita.
  const layoutRef = useRef(layout)
  layoutRef.current = layout

  // O papel precisa das mesmas medidas de faixa que a tela usou, e pela mesma
  // razão da `ref` acima: elas mudam durante a digitação e quem as lê é a
  // impressão, no momento em que ela acontece.
  const bandsRef = useRef(bands)
  bandsRef.current = bands

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
              {/* Os objetos da faixa entram junto: repetem em toda folha,
                  porque a faixa repete. */}
              <FloatingLayer
                objects={[...(floatsByPage[index] ?? []), ...bandFloatsOf(page, index + 1)]}
                page={page}
                schema={editor.schema}
                behind
              />
              {(['header', 'footer'] as const).map((kind) => {
                // A capa manda sobre a paridade, e a paridade sobre o padrão —
                // a ordem do Word. Documento sem primeira página distinta cai no
                // padrão, e nada muda para ele.
                const band = bandForPage(page, index + 1, kind)
                return hasBandContent(band) ? (
                  <PageBand
                    key={kind}
                    band={band}
                    kind={kind}
                    pageNumber={index + 1}
                    totalPages={layout.pages}
                    insetPx={bandInset}
                    offsetPx={mmToPx(kind === 'header' ? page.headerDistanceMm : page.footerDistanceMm)}
                  />
                ) : null
              })}

              <FloatingLayer
                objects={[...(floatsByPage[index] ?? []), ...bandFloatsOf(page, index + 1)]}
                page={page}
                schema={editor.schema}
                behind={false}
              />
            </div>
          ))}

          <div
            className="pages__column"
            style={{
              // A margem de cima é um piso: um cabeçalho mais alto que ela
              // desce o corpo até debaixo dele, como no Word.
              paddingTop: `${mmToPx(insets.top)}px`,
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

/**
 * A altura desenhada do cabeçalho e do rodapé.
 *
 * É a única parte da conta de margem que nenhum arquivo diz: um cabeçalho em
 * grade ocupa o que a fonte e a quebra derem, e isso só existe depois de
 * desenhar. Medido na primeira folha — as outras repetem a mesma faixa — e
 * arredondado a um décimo de milímetro, porque a medida do navegador oscila
 * sozinha e cada oscilação repaginaria o documento inteiro.
 *
 * Não há laço: a altura da faixa não depende de onde o texto caiu.
 */
function useBandHeights(page: PageSetup, revision: number): BandHeights {
  const [bands, setBands] = useState<BandHeights>({ headerMm: 0, footerMm: 0 })

  useEffect(() => {
    const measure = (): void => {
      const sheet = document.querySelector('.paper-bands')
      const heightOf = (kind: string): number => {
        const band = sheet?.querySelector(`.band--${kind}`)
        return band === null || band === undefined
          ? 0
          : Math.round(toMm((band as HTMLElement).offsetHeight) * 10) / 10
      }

      const next = { headerMm: heightOf('header'), footerMm: heightOf('footer') }
      setBands((current) =>
        current.headerMm === next.headerMm && current.footerMm === next.footerMm ? current : next,
      )
    }

    measure()

    const sheet = document.querySelector('.paper-bands')
    if (sheet === null) return undefined

    const observer = new ResizeObserver(measure)
    observer.observe(sheet)
    for (const band of sheet.querySelectorAll('.band')) observer.observe(band)
    return () => observer.disconnect()
  }, [page, revision])

  return bands
}

/**
 * O documento recortado nas folhas que a tela mostra.
 *
 * Serializa **direto dos nós**, e não a partir de `getHTML()`. A diferença não
 * é de gosto: o leitor emite a quebra de página dentro do parágrafo quando o
 * Word a gravou assim (`w:br w:type="page"` no meio de um `w:r`), e um `<div>`
 * dentro de `<p>` faz o analisador de HTML fechar o parágrafo e desalojar o
 * `div`. Um documento de 15 nós virava 17 elementos, os índices deixavam de
 * casar, e o papel cortava em lugar diferente do da tela — que é exatamente o
 * defeito que este trabalho existe para acabar.
 *
 * Serializando o nó, o recorte cai sempre onde o paginador o pôs.
 */
function splitIntoPages(editor: Editor, layout: PageLayout): PrintPage[] {
  const serializer = DOMSerializer.fromSchema(editor.schema)

  const blocks: ProseMirrorNode[] = []
  editor.state.doc.forEach((node: ProseMirrorNode) => blocks.push(node))

  const cortes = [0, ...layout.pageStarts, blocks.length]
  const pages: PrintPage[] = []

  for (let i = 0; i < cortes.length - 1; i++) {
    const inicio = cortes[i]!
    const fim = cortes[i + 1]!
    if (fim <= inicio && pages.length > 0) continue

    const holder = document.createElement('div')
    holder.appendChild(serializer.serializeFragment(Fragment.fromArray(blocks.slice(inicio, fim))))

    // Os objetos ancorados desta folha, com a caixa de texto já serializada: o
    // desenho do papel não conhece o schema do ProseMirror, e quem o conhece é
    // aqui.
    const floats: PrintFloat[] = []
    for (let b = inicio; b < fim; b++) {
      const anchor = layout.anchors[b]
      const block = blocks[b]
      if (anchor === undefined || block === undefined) continue

      for (const object of floatsOf(block.attrs)) {
        floats.push({
          object,
          anchorTopMm: pxToMm(anchor.topPx),
          ...(object.kind === 'text' ? { contentHtml: serializeFloatText(object, editor) } : {}),
        })
      }
    }

    pages.push({ number: pages.length + 1, html: holder.innerHTML, floats })
  }

  return pages
}

/** O conteúdo de uma caixa de texto, em HTML, pelo serializador do editor. */
function serializeFloatText(object: FloatingObject, editor: Editor): string {
  try {
    const nodes = (object.content ?? []).map((node) => ProseMirrorNode.fromJSON(editor.schema, node))
    const holder = document.createElement('div')
    holder.appendChild(DOMSerializer.fromSchema(editor.schema).serializeFragment(Fragment.fromArray(nodes)))
    return holder.innerHTML
  } catch {
    // Caixa que o schema não reconhece sai vazia em vez de derrubar a
    // exportação inteira.
    return ''
  }
}
