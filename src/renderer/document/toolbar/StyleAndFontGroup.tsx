import { useMemo } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { firstFamilyOf } from '@services/document/font-list.js'
import { ToolbarButton, ToolbarGroup, ToolbarSelect } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { focusChain } from './focus-chain.js'
import { blockStyles, FONT_SIZES, withCurrent } from './toolbar-options.js'
import { useFontFamilies } from './useFontFamilies.js'

/** Estilo do bloco, família e tamanho da fonte — o começo da barra. */
export function StyleAndFontGroup({
  editor,
  onOpenStyles,
}: {
  readonly editor: Editor
  readonly onOpenStyles: () => void
}): React.JSX.Element {
  const t = useT()
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      heading: current.isActive('heading') ? String(current.getAttributes('heading')['level'] ?? '') : '',
      // Só o nome da fonte: o que vem do documento é uma pilha de CSS, com a
      // substituta genérica atrás, e é o nome que a lista aqui conhece.
      fontFamily: firstFamilyOf(String(current.getAttributes('textStyle')['fontFamily'] ?? '')),
      fontSize: String(current.getAttributes('textStyle')['fontSize'] ?? '').replace('pt', ''),
    }),
  })

  const styles = useMemo(() => blockStyles(t), [t])
  const fontFamilies = useFontFamilies(active.fontFamily)
  const chain = () => focusChain(editor)

  function applyBlockStyle(value: string): void {
    if (value === 'paragraph') chain().setParagraph().run()
    else
      chain()
        .toggleHeading({ level: Number(value) as 1 | 2 | 3 | 4 })
        .run()
  }

  return (
    <ToolbarGroup label={t('document.styleAndFont.group')}>
      <ToolbarSelect
        label={t('document.styleAndFont.style')}
        value={active.heading === '' ? 'paragraph' : active.heading}
        options={styles}
        onChange={applyBlockStyle}
        width={128}
      />

      <ToolbarSelect
        label={t('document.styleAndFont.font')}
        value={active.fontFamily}
        options={fontFamilies}
        onChange={(value) =>
          value === '' ? chain().unsetFontFamily().run() : chain().setFontFamily(value).run()
        }
        width={150}
      />

      <ToolbarSelect
        label={t('document.styleAndFont.size')}
        value={active.fontSize}
        options={withCurrent(
          [{ value: '', label: '—' }, ...FONT_SIZES.map((size) => ({ value: size, label: size }))],
          active.fontSize,
        )}
        onChange={(value) =>
          value === '' ? chain().unsetFontSize().run() : chain().setFontSize(`${value}pt`).run()
        }
        width={68}
      />

      {/* A lista do documento, ao lado do seletor que aplica os quatro estilos
          que o editor conhece: é aqui que se vê que um `.docx` tem muito mais. */}
      <ToolbarButton icon="styles" label={t('document.styleAndFont.documentStyles')} onClick={onOpenStyles} />
    </ToolbarGroup>
  )
}
