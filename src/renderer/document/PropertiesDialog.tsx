import { useEditorState, type Editor } from '@tiptap/react'
import { FilePropertiesDialog, type StatisticRow } from '../components/FilePropertiesDialog.js'
import { useLanguage } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { tally } from './WordCountDialog.js'
import { updatePropertyFields, type ReferenceContext } from './references.js'

/** The file's properties with the document's counts, as Word shows them. */
export function PropertiesDialog({
  editor,
  context,
  onClose,
}: {
  readonly editor: Editor
  readonly context: () => ReferenceContext
  readonly onClose: () => void
}): React.JSX.Element {
  const language = useLanguage()
  const pages = useWorkspace((state) => state.pageCount)
  const counts = useEditorState({
    editor,
    selector: ({ editor: current }) => tally(current, current.state.doc),
  })
  const number = (value: number): string => value.toLocaleString(language)

  const statistics: readonly StatisticRow[] = [
    ['document.wordCount.pages', number(pages)],
    ['document.wordCount.words', number(counts.words)],
    ['document.wordCount.charactersWithSpaces', number(counts.characters)],
    ['document.wordCount.charactersNoSpaces', number(counts.charactersNoSpaces)],
    ['document.wordCount.paragraphs', number(counts.paragraphs)],
  ]

  return (
    <FilePropertiesDialog
      title="document.properties.title.dialog"
      statistics={statistics}
      onApply={() => updatePropertyFields(editor, context())}
      onClose={() => {
        onClose()
        requestAnimationFrame(() => editor.view.focus())
      }}
    />
  )
}
