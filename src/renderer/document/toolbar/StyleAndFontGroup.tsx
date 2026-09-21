import { useEditorState, type Editor } from '@tiptap/react'
import { firstFamilyOf } from '@services/document/font-list.js'
import { ToolbarGroup, ToolbarSelect } from '../../components/ToolbarControls.js'
import { focusChain } from './focus-chain.js'
import { BLOCK_STYLES, FONT_SIZES, withCurrent } from './toolbar-options.js'
import { useFontFamilies } from './useFontFamilies.js'

/** Estilo do bloco, família e tamanho da fonte — o começo da barra. */
export function StyleAndFontGroup({ editor }: { readonly editor: Editor }): React.JSX.Element {
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
    <ToolbarGroup label="Estilos e fonte">
      <ToolbarSelect
        label="Estilo"
        value={active.heading === '' ? 'paragraph' : active.heading}
        options={BLOCK_STYLES}
        onChange={applyBlockStyle}
        width={128}
      />

      <ToolbarSelect
        label="Fonte"
        value={active.fontFamily}
        options={fontFamilies}
        onChange={(value) =>
          value === '' ? chain().unsetFontFamily().run() : chain().setFontFamily(value).run()
        }
        width={150}
      />

      <ToolbarSelect
        label="Tamanho"
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
    </ToolbarGroup>
  )
}
