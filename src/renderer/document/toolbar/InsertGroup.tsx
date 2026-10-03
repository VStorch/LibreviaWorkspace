import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'
import { appendPageField } from '@services/document/band.js'
import { focusChain } from './focus-chain.js'

interface InsertGroupProps {
  readonly editor: Editor
  readonly onOpenLink: () => void
  /** O menu "Tabela" e o de contexto chegam aos mesmos diálogos, abertos de fora. */
  readonly onOpenTable: () => void
  readonly onOpenImageProperties: () => void
}

export function InsertGroup({
  editor,
  onOpenLink,
  onOpenTable,
  onOpenImageProperties,
}: InsertGroupProps): React.JSX.Element {
  const t = useT()
  const showError = useWorkspace((state) => state.showError)

  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      link: current.isActive('link'),
      onImage: current.isActive('image'),
    }),
  })

  const chain = () => focusChain(editor)

  const readOnly = useWorkspace((state) => state.readOnly)

  /** Na peça da faixa com o cursor, como `{n}`, que a gravação transforma em campo; senão, no fim do rodapé. */
  function insertPageField(token: '{n}' | '{total}'): void {
    if (readOnly) return
    const focused = document.activeElement
    if (focused instanceof HTMLElement && focused.classList.contains('band__text')) {
      document.execCommand('insertText', false, token)
      return
    }
    const { page, setPage } = useWorkspace.getState()
    const updated = appendPageField(page, token)
    if (updated !== null) setPage(updated)
  }

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
    <ToolbarGroup label={t('document.insert.group')}>
      {/* Excluir a tabela mora no menu "Tabela" e no botão direito. */}
      <ToolbarButton icon="table" label={t('document.insert.table')} onClick={onOpenTable} />
      {/* Com uma imagem selecionada, abre as propriedades dela. */}
      <ToolbarButton
        icon="image"
        label={active.onImage ? t('document.imageDialog.title') : t('document.insert.image')}
        active={active.onImage}
        onClick={() => (active.onImage ? onOpenImageProperties() : void insertImage())}
      />
      <ToolbarButton
        icon="link"
        label={t('document.insert.link')}
        active={active.link}
        onClick={onOpenLink}
      />
      <ToolbarButton
        icon="page-break"
        label={t('menu.insert.pageBreak')}
        shortcut={shortcutHintOf(SHORTCUTS.insertPageBreak)}
        onClick={() => chain().setPageBreak().run()}
      />
      <ToolbarButton
        icon="page-number"
        label={t('document.insert.pageNumber')}
        onClick={() => insertPageField('{n}')}
      />
      <ToolbarButton
        icon="page-total"
        label={t('document.insert.totalPages')}
        onClick={() => insertPageField('{total}')}
      />
    </ToolbarGroup>
  )
}
