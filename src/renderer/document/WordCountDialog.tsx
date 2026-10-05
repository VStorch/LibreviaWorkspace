import { useEditorState, type Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { DocumentNode } from '@services/document/model.js'
import { charactersWithoutSpaces, countParagraphs } from '@services/document/word-count.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { textWithoutDeletions } from './extensions/track-changes.js'

/**
 * Two columns, document and selection: the question is usually about **this passage**. Words and
 * characters come from `CharacterCount`, as in the status bar.
 */
export function WordCountDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const pages = useWorkspace((state) => state.pageCount)

  // Live: the text may change with the dialog open.
  const counts = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const { from, to, empty } = current.state.selection
      return {
        document: tally(current, current.state.doc),
        // `doc.cut` returns the node `CharacterCount` accepts.
        selection: empty ? null : tally(current, current.state.doc.cut(from, to)),
      }
    },
  })

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('document.wordCount.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <table className="counts">
        <thead>
          <tr>
            <th scope="col">{t('document.wordCount.count')}</th>
            <th scope="col">{t('document.wordCount.document')}</th>
            <th scope="col">{t('document.wordCount.selection')}</th>
          </tr>
        </thead>
        <tbody>
          <Row
            label={t('document.wordCount.words')}
            document={counts.document.words}
            selection={counts.selection?.words}
          />
          <Row
            label={t('document.wordCount.charactersWithSpaces')}
            document={counts.document.characters}
            selection={counts.selection?.characters}
          />
          <Row
            label={t('document.wordCount.charactersNoSpaces')}
            document={counts.document.charactersNoSpaces}
            selection={counts.selection?.charactersNoSpaces}
          />
          <Row
            label={t('document.wordCount.paragraphs')}
            document={counts.document.paragraphs}
            selection={counts.selection?.paragraphs}
          />
          {/* Pages only for the document: "half a page selected" would be an invented number. */}
          <Row label={t('document.wordCount.pages')} document={pages} selection={undefined} />
        </tbody>
      </table>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" autoFocus onClick={onClose}>
          {t('document.common.close')}
        </button>
      </div>
    </div>
  )
}

export interface Tally {
  readonly words: number
  readonly characters: number
  readonly charactersNoSpaces: number
  readonly paragraphs: number
}

export function tally(editor: Editor, node: ProseMirrorNode): Tally {
  const storage = editor.storage['characterCount']

  return {
    words: storage.words({ node }),
    characters: storage.characters({ node }),
    // The same text `CharacterCount` measures, so "with" minus "without" spaces gives the spaces.
    charactersNoSpaces: charactersWithoutSpaces(textWithoutDeletions(node, undefined, ' ')),
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
      {/* "Nothing selected" is not "zero words". */}
      <td>{selection === undefined ? '—' : selection.toLocaleString('pt-BR')}</td>
    </tr>
  )
}
