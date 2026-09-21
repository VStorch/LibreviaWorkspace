import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { focusChain } from './focus-chain.js'

/** Listas e recuo: o que muda a estrutura do bloco, e não a aparência do texto. */
export function ListAndIndentGroup({ editor }: { readonly editor: Editor }): React.JSX.Element {
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      bulletList: current.isActive('bulletList'),
      orderedList: current.isActive('orderedList'),
    }),
  })

  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label="Listas e recuo">
      <ToolbarButton
        icon="bullet-list"
        label="Lista com marcadores"
        active={active.bulletList}
        onClick={() => chain().toggleBulletList().run()}
      />
      <ToolbarButton
        icon="ordered-list"
        label="Lista numerada"
        active={active.orderedList}
        onClick={() => chain().toggleOrderedList().run()}
      />
      <ToolbarButton
        icon="outdent"
        label="Diminuir recuo"
        shortcut={shortcutHintOf(SHORTCUTS.outdent)}
        onClick={() => chain().outdent().run()}
      />
      <ToolbarButton
        icon="indent"
        label="Aumentar recuo"
        shortcut={shortcutHintOf(SHORTCUTS.indent)}
        onClick={() => chain().indent().run()}
      />
    </ToolbarGroup>
  )
}
