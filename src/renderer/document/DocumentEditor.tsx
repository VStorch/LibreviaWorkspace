import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { JSONContent } from '@tiptap/core'
import { EditorContent, useEditor } from '@tiptap/react'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import type { ContextMenuTarget } from '@shared/types.js'
import { DOCUMENT_CONTENT_CSS, EDITOR_ONLY_CSS } from '@services/document/content-styles.js'
import { plainPasteContent } from '@services/document/paste.js'
import {
  contentInsetsMm,
  mmToPx,
  pageDimensionsMm,
  pxToMm,
  type DocumentNode,
} from '@services/document/model.js'
import { editBandFloat, editBandPiece } from '@services/document/band.js'
import { floatsOf } from '@services/document/floating.js'
import { currentPreferences, usePreferences } from '../state/preferences.js'
import { useLeaveReadingOnEscape, useReadingMode } from '../state/reading.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { TableDialog } from './toolbar/TableDialog.js'
import { TablePropertiesDialog } from './toolbar/TablePropertiesDialog.js'
import { ImageDialog } from './toolbar/ImageDialog.js'
import { DocumentContextMenu } from './DocumentContextMenu.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { SpecialCharsDialog } from './SpecialCharsDialog.js'
import { StylesPanel } from './StylesPanel.js'
import { WordCountDialog } from './WordCountDialog.js'
import { PaperSheet } from './PaperSheet.js'
import { usePagination } from './usePagination.js'
import { useBandHeights } from './useBandHeights.js'
import { splitIntoPages } from './print-source.js'
import type { FloatSource, PlacedFloat } from './FloatingLayer.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { isPaginationOnly } from './extensions/pagination.js'
import { useEditorCommands } from './useEditorCommands.js'
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
  const setPage = useWorkspace((state) => state.setPage)
  const showError = useWorkspace((state) => state.showError)
  const preferences = usePreferences((state) => state.preferences)
  const reading = useReadingMode()

  useLeaveReadingOnEscape(reading)

  const pageRef = useRef<HTMLDivElement>(null)
  const [contentRevision, setContentRevision] = useState(0)

  const [searchStatus, setSearchStatus] = useState<SearchStatus>({ total: 0, current: 0 })
  const [contextTarget, setContextTarget] = useState<ContextMenuTarget | null>(null)

  const handleSearchStatus = useCallback((status: SearchStatus) => setSearchStatus(status), [])

  const editor = useEditor({
    // As preferências são lidas da loja, e não das props: o editor é criado uma
    // vez só, e o que muda depois chega pelos efeitos mais abaixo.
    extensions: buildEditorExtensions(handleSearchStatus, {
      isTypographyEnabled: () => currentPreferences().typography,
      invisibleCharactersVisible: currentPreferences().invisibleCharacters,
    }),
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
      attributes: {
        class: 'page__content',
        spellcheck: currentPreferences().spellcheck ? 'true' : 'false',
      },
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
  // O modo de leitura trava a edicao junto com o somente leitura, e pelo
  // mesmo caminho: uma tecla perdida nao pode alterar o documento que a
  // pessoa esta lendo. Sair do modo devolve a edicao sem recriar o editor,
  // entao o cursor e o historico sobrevivem a ida e volta.
  useEffect(() => {
    editor?.setEditable(!readOnly && !reading, false)
  }, [editor, readOnly, reading])

  // Salvar e imprimir precisam do conteúdo atual, que só o editor conhece.
  useEffect(() => {
    if (editor === null) return undefined
    registerDocumentSource({
      readDoc: () => editor.getJSON() as DocumentNode,
      readHtml: () => editor.getHTML(),
      readPages: () => ({
        pages: splitIntoPages(editor, layoutRef.current, page),
        bands: bandsRef.current,
      }),
    })
    return () => registerDocumentSource(null)
  }, [editor, page, registerDocumentSource])

  /**
   * Ortografia ligada e desligada no editor já montado.
   *
   * O atributo é escrito no elemento, e não passado de novo em `editorProps`: o
   * ProseMirror só lê os atributos ao criar a visão, e recriá-la aqui perderia o
   * cursor e o histórico — o mesmo motivo do `setEditable` acima.
   *
   * Quem de fato liga o corretor é o processo main, na sessão do Chromium. Este
   * atributo é a outra metade: sem ele, o campo continua marcado como "não
   * verifique".
   */
  useEffect(() => {
    editor?.view.dom.setAttribute('spellcheck', preferences.spellcheck ? 'true' : 'false')
  }, [editor, preferences.spellcheck])

  // As marcas de formatação são um comando, e a transação dele não muda o
  // documento — então não suja o arquivo nem dispara nova medição.
  useEffect(() => {
    editor?.commands.showInvisibleCharacters(preferences.invisibleCharacters)
  }, [editor, preferences.invisibleCharacters])

  /**
   * Colar sem formatação.
   *
   * O texto vem do main (só ele alcança a área de transferência) e a conversão em
   * parágrafos é função pura, testada em `@services/document/paste.ts`. Nada de
   * `pasteAndMatchStyle` do Chromium: ele **adapta** a formatação em vez de
   * descartá-la, e um trecho colado de uma página da web chegava com tamanho de
   * fonte e cor próprios.
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

    // O elenco existe porque `DocumentNode` é o nosso modelo e `JSONContent` é o
    // do Tiptap: as duas formas são a mesma, e é o serviço puro que a garante.
    editor
      .chain()
      .focus()
      .insertContent(content as JSONContent[])
      .run()
  }, [editor, readOnly, showError])

  // O menu nativo e o botão direito chegam pelo mesmo `run`, que é onde a trava
  // do somente leitura é conferida — ver useEditorCommands.
  const { dialogs, setDialog, run } = useEditorCommands(editor, readOnly, pasteWithoutFormat)

  // O botão direito nasce no processo main: é lá que o corretor do Chromium conta
  // qual palavra marcou e o que sugere.
  useEffect(
    () =>
      window.api.contextMenu.onRequest((payload) => {
        // Validado com o mesmo contrato que o main usou para mandar — a segunda
        // ponta do zod, que o preload não pode fazer por rodar sandboxed.
        const parsed = pushContracts[IpcChannel.ContextMenuRequested].safeParse(payload)
        if (!parsed.success) return

        // Fora de campo editável e sem nada selecionado não há ação a oferecer: o
        // menu apareceria com todos os itens apagados, que é pior que menu nenhum.
        // É o caso do clique na barra de ferramentas e na barra de status.
        if (!parsed.data.editable && !parsed.data.canCopy) return

        setContextTarget(parsed.data)
      }),
    [],
  )

  const bands = useBandHeights(page, contentRevision)
  const layout = usePagination(editor, page, contentRevision, bands, !reading)
  const insets = contentInsetsMm(page, bands)

  // Os objetos ancorados de cada folha. Recalculados junto com a paginação
  // porque a posição de um deles depende de em que folha o parágrafo âncora
  // caiu — e isso muda a cada linha digitada.
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
        // A posição do bloco viaja junto: é por ela que o texto digitado dentro
        // da caixa acha o caminho de volta ao atributo de onde saiu.
        sheet.push({ object, anchorTopMm: pxToMm(anchor.topPx), source: { pos, index: slot } })
        slot += 1
      }
    })

    return pages
  }, [editor, layout, contentRevision])

  /**
   * O texto digitado dentro de uma caixa volta para o atributo do bloco.
   *
   * Uma transação comum, e não um caminho paralelo: assim a edição entra no
   * histórico, marca o documento como alterado e chega ao gravador pelo mesmo
   * `getJSON()` de todo o resto.
   */
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

  /**
   * O texto digitado no cabeçalho ou no rodapé volta para a configuração.
   *
   * A faixa não mora no documento do editor — ela é a parte OOXML preservada, e
   * vive em `page`. Por isso a volta é `setPage` e não uma transação: o
   * histórico do editor não tem o que desfazer aqui, e o gravador lê a
   * configuração pelo mesmo caminho de sempre.
   */
  const editBand = useCallback(
    (pid: string, text: string) => {
      if (readOnly) return

      // A configuração vem da loja e não da renderização: várias peças podem
      // sair do foco em sequência, e uma leitura presa no fechamento apagaria
      // a edição anterior a cada uma delas.
      const current = useWorkspace.getState().page
      const updated = editBandPiece(current, pid, text)
      if (updated !== current) setPage(updated)
    },
    [readOnly, setPage],
  )

  /**
   * O texto digitado numa caixa da faixa volta para a configuração.
   *
   * A caixa vem inteira, e não parágrafo a parágrafo: digitar dentro dela abre
   * e fecha parágrafos, e um endereço por parágrafo quebraria no primeiro Enter.
   */
  const editBandBox = useCallback(
    (bid: string, content: DocumentNode[]) => {
      if (readOnly) return

      const current = useWorkspace.getState().page
      const updated = editBandFloat(current, bid, content)
      if (updated !== current) setPage(updated)
    },
    [readOnly, setPage],
  )

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
  const editableSheet = readOnly
    ? {}
    : { onEditFloat: editFloat, onEditBandPiece: editBand, onEditBandBox: editBandBox }

  return (
    <div className="editor-shell">
      {/* O estilo do conteúdo vem do mesmo módulo que o HTML de impressão usa.
          Duas folhas de estilo divergiriam com o tempo, e o PDF deixaria de
          sair igual à tela — o risco registrado no §6.3 do plano. */}
      <style>{DOCUMENT_CONTENT_CSS + EDITOR_ONLY_CSS}</style>

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
        />
      )}

      {dialogs.find && (
        <FindReplacePanel editor={editor} status={searchStatus} onClose={() => setDialog('find', false)} />
      )}

      {dialogs.pageSetup && <PageSetupPanel onClose={() => setDialog('pageSetup', false)} />}

      {dialogs.wordCount && <WordCountDialog editor={editor} onClose={() => setDialog('wordCount', false)} />}

      {dialogs.styles && <StylesPanel editor={editor} onClose={() => setDialog('styles', false)} />}

      {dialogs.specialCharacter && (
        <SpecialCharsDialog editor={editor} onClose={() => setDialog('specialCharacter', false)} />
      )}

      {dialogs.table && <TableDialog editor={editor} onClose={() => setDialog('table', false)} />}

      {dialogs.tableProperties && (
        <TablePropertiesDialog editor={editor} onClose={() => setDialog('tableProperties', false)} />
      )}

      {dialogs.imageProperties && (
        <ImageDialog editor={editor} onClose={() => setDialog('imageProperties', false)} />
      )}

      {contextTarget !== null && (
        <DocumentContextMenu
          target={contextTarget}
          // O menu de contexto só oferece as ações de tabela quando o cursor está
          // dentro de uma: fora dela, "mesclar células" não tem o que mesclar.
          inTable={editor.isActive('table')}
          onTableAction={run}
          onClose={() => setContextTarget(null)}
          onPasteWithoutFormat={() => void pasteWithoutFormat()}
        />
      )}

      <div className={`editor-scroll${reading ? ' editor-scroll--reading' : ''}`}>
        <div
          ref={pageRef}
          className={`pages${reading ? ' pages--reading' : ''}`}
          style={reading ? undefined : { width: `${mmToPx(width)}px`, height: `${layout.stackHeightPx}px` }}
        >
          {/* As folhas: papel desenhado atrás do texto. Ficam fora do
              `contenteditable` de propósito — dentro dele, cada folha seria um
              nó que a pessoa conseguiria selecionar e apagar.

              No modo de leitura não há folha nenhuma: a pilha de papel é o que
              a rolagem contínua existe para tirar da frente. Os objetos
              ancorados saem junto, e não por descuido — a posição deles é
              relativa a uma folha, e sem folha não há onde pousá-los. */}
          {!reading &&
            layout.sheetTops.map((top, index) => (
              <div
                key={top}
                className={`paper${(layout.sheetHeights[index] ?? 0) > mmToPx(height) + 1 ? ' paper--oversized' : ''}`}
                style={{ top: `${top}px`, height: `${layout.sheetHeights[index] ?? mmToPx(height)}px` }}
                aria-hidden="true"
              >
                <span className="paper__number">{index + 1}</span>
              </div>
            ))}

          {/* Uma faixa por folha, com o número real. No papel elas moram dentro
              da margem, e é por isso que não empurram o texto. Sem folhas não
              há cabeçalho repetido: "página 3 de 12" não quer dizer nada numa
              tira contínua. */}
          {!reading &&
            layout.sheetTops.map((top, index) => (
              <PaperSheet
                key={`banda-${top}`}
                page={page}
                pageNumber={index + 1}
                totalPages={layout.pages}
                topPx={top}
                floats={floatsByPage[index] ?? []}
                schema={editor.schema}
                {...editableSheet}
              />
            ))}

          <div
            className="pages__column"
            style={
              reading
                ? // A largura da leitura vem do CSS e nao das margens do
                  // documento: uma margem de 10 mm daria uma linha larga
                  // demais para ler com conforto, e o modo existe justamente
                  // para nao obedecer ao papel.
                  undefined
                : {
                    // A margem de cima é um piso: um cabeçalho mais alto que ela
                    // desce o corpo até debaixo dele, como no Word.
                    paddingTop: `${mmToPx(insets.top)}px`,
                    paddingRight: `${mmToPx(page.margins.right)}px`,
                    paddingLeft: `${mmToPx(page.margins.left)}px`,
                  }
            }
          >
            <EditorContent editor={editor} />
          </div>
        </div>
      </div>
    </div>
  )
}
