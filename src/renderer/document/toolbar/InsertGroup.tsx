import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useWorkspace } from '../../state/workspace.js'
import { focusChain } from './focus-chain.js'

interface InsertGroupProps {
  readonly editor: Editor
  /** O diálogo do link é desenhado pela barra, junto com os outros. */
  readonly onOpenLink: () => void
}

/** O que se põe dentro do documento: tabela, imagem, link, quebra de página. */
export function InsertGroup({ editor, onOpenLink }: InsertGroupProps): React.JSX.Element {
  const showError = useWorkspace((state) => state.showError)

  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      inTable: current.isActive('table'),
      link: current.isActive('link'),
    }),
  })

  const chain = () => focusChain(editor)

  async function insertImage(): Promise<void> {
    const result = await window.api.image.pick({})
    if (!result.ok) {
      showError(result.error)
      return
    }
    if (result.data.canceled) return
    chain().setImage({ src: result.data.dataUrl, alt: result.data.name }).run()
  }

  return (
    <ToolbarGroup label="Inserir">
      <ToolbarButton
        icon="table"
        label={active.inTable ? 'Remover tabela' : 'Inserir tabela'}
        active={active.inTable}
        onClick={() =>
          active.inTable
            ? chain().deleteTable().run()
            : chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
        }
      />
      <ToolbarButton icon="image" label="Inserir imagem" onClick={() => void insertImage()} />
      <ToolbarButton icon="link" label="Inserir link" active={active.link} onClick={onOpenLink} />
      <ToolbarButton
        icon="page-break"
        label="Quebra de página"
        shortcut={shortcutHintOf(SHORTCUTS.insertPageBreak)}
        onClick={() => chain().setPageBreak().run()}
      />
    </ToolbarGroup>
  )
}
