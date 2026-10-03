import { useMemo } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup, ToolbarSelect } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { blockLineHeightOf } from '../extensions/paragraph-commands.js'
import { focusChain } from './focus-chain.js'
import { lineHeights, withCurrent } from './toolbar-options.js'

interface ParagraphGroupProps {
  readonly editor: Editor
  /** O menu nativo também abre o parágrafo. */
  readonly paragraphOpen: boolean
  readonly onParagraphOpenChange: (open: boolean) => void
}

export function ParagraphGroup({
  editor,
  paragraphOpen,
  onParagraphOpenChange,
}: ParagraphGroupProps): React.JSX.Element {
  const t = useT()
  const lineHeightOptions = useMemo(() => lineHeights(t), [t])
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      alignLeft: current.isActive({ textAlign: 'left' }),
      alignCenter: current.isActive({ textAlign: 'center' }),
      alignRight: current.isActive({ textAlign: 'right' }),
      alignJustify: current.isActive({ textAlign: 'justify' }),
      // Do **bloco**, e não da marca de texto: no OOXML não existe `w:line` num `w:rPr`.
      lineHeight: blockLineHeightOf(current),
    }),
  })

  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label={t('document.paragraph.title')}>
      <ToolbarButton
        icon="align-left"
        label={t('document.paragraph.alignLeftLabel')}
        shortcut={shortcutHintOf(SHORTCUTS.alignLeft)}
        active={active.alignLeft}
        onClick={() => chain().setTextAlign('left').run()}
      />
      <ToolbarButton
        icon="align-center"
        label={t('document.paragraph.alignCenterLabel')}
        shortcut={shortcutHintOf(SHORTCUTS.alignCenter)}
        active={active.alignCenter}
        onClick={() => chain().setTextAlign('center').run()}
      />
      <ToolbarButton
        icon="align-right"
        label={t('document.paragraph.alignRightLabel')}
        shortcut={shortcutHintOf(SHORTCUTS.alignRight)}
        active={active.alignRight}
        onClick={() => chain().setTextAlign('right').run()}
      />
      <ToolbarButton
        icon="align-justify"
        label={t('document.paragraph.alignJustifyLabel')}
        shortcut={shortcutHintOf(SHORTCUTS.alignJustify)}
        active={active.alignJustify}
        onClick={() => chain().setTextAlign('justify').run()}
      />

      <ToolbarSelect
        label={t('document.paragraph.lineSpacingLabel')}
        value={active.lineHeight}
        // O atributo guarda `1.5`, e a tela escreve 1,5.
        options={withCurrent(lineHeightOptions, active.lineHeight, (value) => value.replace('.', ','))}
        // A conversão para o CSS é bloco a bloco, porque depende da fonte.
        onChange={(value) => chain().setBlockLineHeight(value).run()}
        width={100}
      />

      <ToolbarButton
        icon="paragraph"
        label={t('menu.format.paragraph')}
        active={paragraphOpen}
        onClick={() => onParagraphOpenChange(!paragraphOpen)}
      />
    </ToolbarGroup>
  )
}
