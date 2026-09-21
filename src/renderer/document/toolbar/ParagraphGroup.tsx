import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup, ToolbarSelect } from '../../components/ToolbarControls.js'
import { blockLineHeightOf } from '../extensions/paragraph-commands.js'
import { focusChain } from './focus-chain.js'
import { LINE_HEIGHTS, withCurrent } from './toolbar-options.js'

interface ParagraphGroupProps {
  readonly editor: Editor
  /** Aberto de fora também: o menu nativo tem "Formatar → Parágrafo…". */
  readonly paragraphOpen: boolean
  readonly onParagraphOpenChange: (open: boolean) => void
}

/** O que é propriedade do parágrafo: alinhamento, entrelinha e o diálogo completo. */
export function ParagraphGroup({
  editor,
  paragraphOpen,
  onParagraphOpenChange,
}: ParagraphGroupProps): React.JSX.Element {
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      alignLeft: current.isActive({ textAlign: 'left' }),
      alignCenter: current.isActive({ textAlign: 'center' }),
      alignRight: current.isActive({ textAlign: 'right' }),
      alignJustify: current.isActive({ textAlign: 'justify' }),
      // Do **bloco**, e não da marca de texto: no OOXML a entrelinha é
      // propriedade do parágrafo, e não existe `w:line` dentro de um `w:rPr`.
      // Enquanto este seletor escrevia na marca, escolher "Duplo" aqui era perda
      // garantida — o gravador não tinha onde pôr a medida e a anotava no
      // inventário.
      lineHeight: blockLineHeightOf(current),
    }),
  })

  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label="Parágrafo">
      <ToolbarButton
        icon="align-left"
        label="Alinhar à esquerda"
        shortcut={shortcutHintOf(SHORTCUTS.alignLeft)}
        active={active.alignLeft}
        onClick={() => chain().setTextAlign('left').run()}
      />
      <ToolbarButton
        icon="align-center"
        label="Centralizar"
        shortcut={shortcutHintOf(SHORTCUTS.alignCenter)}
        active={active.alignCenter}
        onClick={() => chain().setTextAlign('center').run()}
      />
      <ToolbarButton
        icon="align-right"
        label="Alinhar à direita"
        shortcut={shortcutHintOf(SHORTCUTS.alignRight)}
        active={active.alignRight}
        onClick={() => chain().setTextAlign('right').run()}
      />
      <ToolbarButton
        icon="align-justify"
        label="Justificar"
        shortcut={shortcutHintOf(SHORTCUTS.alignJustify)}
        active={active.alignJustify}
        onClick={() => chain().setTextAlign('justify').run()}
      />

      <ToolbarSelect
        label="Espaçamento entre linhas"
        value={active.lineHeight}
        // A vírgula é a nossa: o atributo guarda `1.5`, e a tela escreve 1,5.
        options={withCurrent(LINE_HEIGHTS, active.lineHeight, (value) => value.replace('.', ','))}
        // O valor é a escolha em linhas — vazio é "Simples" —, e a conversão para
        // a medida do CSS acontece bloco a bloco, porque depende da fonte.
        onChange={(value) => chain().setBlockLineHeight(value).run()}
        width={100}
      />

      <ToolbarButton
        icon="paragraph"
        label="Parágrafo…"
        active={paragraphOpen}
        onClick={() => onParagraphOpenChange(!paragraphOpen)}
      />
    </ToolbarGroup>
  )
}
