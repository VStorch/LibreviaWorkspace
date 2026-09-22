import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { focusChain } from './focus-chain.js'

/** Listas e recuo: o que muda a estrutura do bloco, e não a aparência do texto. */
export function ListAndIndentGroup({ editor }: { readonly editor: Editor }): React.JSX.Element {
  const t = useT()
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      bulletList: current.isActive('bulletList'),
      orderedList: current.isActive('orderedList'),
    }),
  })

  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label={t('document.listAndIndent.group')}>
      <ToolbarButton
        icon="bullet-list"
        label={t('document.listAndIndent.bulletList')}
        active={active.bulletList}
        onClick={() => chain().toggleBulletList().run()}
      />
      <ToolbarButton
        icon="ordered-list"
        label={t('document.listAndIndent.numberedList')}
        active={active.orderedList}
        onClick={() => chain().toggleOrderedList().run()}
      />
      <ToolbarButton
        icon="outdent"
        label={t('document.listAndIndent.decreaseIndent')}
        shortcut={shortcutHintOf(SHORTCUTS.outdent)}
        onClick={() => chain().outdent().run()}
      />
      <ToolbarButton
        icon="indent"
        label={t('document.listAndIndent.increaseIndent')}
        shortcut={shortcutHintOf(SHORTCUTS.indent)}
        onClick={() => chain().indent().run()}
      />
    </ToolbarGroup>
  )
}
