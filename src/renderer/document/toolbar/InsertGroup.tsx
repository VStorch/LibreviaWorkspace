import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useWorkspace } from '../../state/workspace.js'
import { focusChain } from './focus-chain.js'

interface InsertGroupProps {
  readonly editor: Editor
  /** O diálogo do link é desenhado pela barra, junto com os outros. */
  readonly onOpenLink: () => void
  /**
   * Os dois abertos de fora: o menu nativo "Tabela" e o menu de contexto chegam
   * aos mesmos diálogos, e um diálogo por caminho de abertura seriam dois.
   */
  readonly onOpenTable: () => void
  readonly onOpenImageProperties: () => void
}

/** O que se põe dentro do documento: tabela, imagem, link, quebra de página. */
export function InsertGroup({
  editor,
  onOpenLink,
  onOpenTable,
  onOpenImageProperties,
}: InsertGroupProps): React.JSX.Element {
  const showError = useWorkspace((state) => state.showError)

  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      link: current.isActive('link'),
      onImage: current.isActive('image'),
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
      {/* Abre o diálogo em vez de inserir uma 3 × 3 fixa: quem precisa de cinco
          colunas não tem de acrescentá-las uma a uma depois. Excluir a tabela
          mora no menu "Tabela" e no botão direito, com o resto da estrutura. */}
      <ToolbarButton icon="table" label="Inserir tabela" onClick={onOpenTable} />
      {/* Com uma imagem selecionada o botão passa a abrir as propriedades dela —
          texto alternativo e alinhamento —, que é onde se espera procurá-las. */}
      <ToolbarButton
        icon="image"
        label={active.onImage ? 'Propriedades da imagem' : 'Inserir imagem'}
        active={active.onImage}
        onClick={() => (active.onImage ? onOpenImageProperties() : void insertImage())}
      />
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
