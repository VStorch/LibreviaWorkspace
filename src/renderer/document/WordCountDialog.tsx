import { useEditorState, type Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { DocumentNode } from '@services/document/model.js'
import { charactersWithoutSpaces, countParagraphs } from '@services/document/word-count.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Contagem de palavras — o diálogo do Word.
 *
 * Duas colunas, documento e seleção, porque a pergunta real quase nunca é
 * "quantas palavras tem o arquivo": é "quantas palavras tem **este trecho**", com
 * um limite de laudas ou de caracteres para cumprir.
 *
 * Palavras e caracteres vêm da extensão `CharacterCount`, a mesma que alimenta a
 * barra de status. Contá-los aqui de outro jeito daria dois números para a mesma
 * coisa na mesma tela.
 */
export function WordCountDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const pages = useWorkspace((state) => state.estimatedPages)

  // Ao vivo: enquanto o diálogo está aberto o texto pode mudar — e muda, porque
  // fechar o diálogo não é condição para continuar escrevendo.
  const counts = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const { from, to, empty } = current.state.selection
      return {
        document: tally(current, current.state.doc),
        // `doc.cut` devolve um nó de verdade, e é ele que o `CharacterCount`
        // aceita: assim a seleção é contada pela mesma régua do documento.
        selection: empty ? null : tally(current, current.state.doc.cut(from, to)),
      }
    },
  })

  return (
    <div
      className="popover"
      role="dialog"
      aria-label="Contagem de palavras"
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <table className="counts">
        <thead>
          <tr>
            <th scope="col">Contagem</th>
            <th scope="col">Documento</th>
            <th scope="col">Seleção</th>
          </tr>
        </thead>
        <tbody>
          <Row label="Palavras" document={counts.document.words} selection={counts.selection?.words} />
          <Row
            label="Caracteres (com espaços)"
            document={counts.document.characters}
            selection={counts.selection?.characters}
          />
          <Row
            label="Caracteres (sem espaços)"
            document={counts.document.charactersNoSpaces}
            selection={counts.selection?.charactersNoSpaces}
          />
          <Row
            label="Parágrafos"
            document={counts.document.paragraphs}
            selection={counts.selection?.paragraphs}
          />
          {/* Páginas só do documento: a folha em que um trecho cai é a mesma
              informação que a barra de status já dá, e "meia página selecionada"
              seria um número inventado. */}
          <Row label="Páginas" document={pages} selection={undefined} />
        </tbody>
      </table>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" autoFocus onClick={onClose}>
          Fechar
        </button>
      </div>
    </div>
  )
}

interface Tally {
  readonly words: number
  readonly characters: number
  readonly charactersNoSpaces: number
  readonly paragraphs: number
}

function tally(editor: Editor, node: ProseMirrorNode): Tally {
  // A mesma extensão que a barra de status usa. Ela aceita um nó, e é isso que
  // permite contar a seleção sem escrever um segundo contador.
  const storage = editor.storage['characterCount']

  return {
    words: storage.words({ node }),
    characters: storage.characters({ node }),
    // Pelo mesmo texto que o `CharacterCount` mede — separador de bloco nenhum,
    // nó folha como um espaço — para que "com espaços" menos "sem espaços" seja
    // exatamente a quantidade de espaços.
    charactersNoSpaces: charactersWithoutSpaces(node.textBetween(0, node.content.size, undefined, ' ')),
    paragraphs: countParagraphs(node.toJSON() as DocumentNode),
  }
}

function Row({
  label,
  document,
  selection,
}: {
  readonly label: string
  readonly document: number
  readonly selection: number | undefined
}): React.JSX.Element {
  return (
    <tr>
      <th scope="row">{label}</th>
      <td>{document.toLocaleString('pt-BR')}</td>
      {/* Travessão, e não zero: "nada selecionado" e "seleção com zero palavras"
          são coisas diferentes. */}
      <td>{selection === undefined ? '—' : selection.toLocaleString('pt-BR')}</td>
    </tr>
  )
}
