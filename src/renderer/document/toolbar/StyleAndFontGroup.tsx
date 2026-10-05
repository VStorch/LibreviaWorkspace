import { useMemo } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { firstFamilyOf } from '@services/document/font-list.js'
import { StyleType, blockStyleOf, listedStyles, styleLabelOf } from '@services/document/styles.js'
import { ToolbarButton, ToolbarGroup, ToolbarSelect } from '../../components/ToolbarControls.js'
import { useLanguage, useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'
import { focusChain } from './focus-chain.js'
import { FONT_SIZES, withCurrent } from './toolbar-options.js'
import { useFontFamilies } from './useFontFamilies.js'

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
      block: (() => {
        const parent = current.state.selection.$from.parent
        const styleId = parent.attrs['styleId']
        const level = parent.attrs['level']
        return {
          type: parent.type.name,
          styleId: typeof styleId === 'string' ? styleId : null,
          level: typeof level === 'number' ? level : null,
        }
      })(),
      // The name, not the CSS stack the document carries.
      fontFamily: firstFamilyOf(String(current.getAttributes('textStyle')['fontFamily'] ?? '')),
      fontSize: String(current.getAttributes('textStyle')['fontSize'] ?? '').replace('pt', ''),
    }),
  })

  const sheet = useWorkspace((state) => state.styles)
  const language = useLanguage()
  // The **document's** styles, by screen name; heading ↔ paragraph by the name `heading N`.
  const styles = useMemo(
    () =>
      listedStyles(sheet, language)
        .filter((style) => style.type === StyleType.Paragraph)
        .map((style) => ({ value: style.id, label: styleLabelOf(style, language) })),
    [sheet, language],
  )
  const currentStyle = blockStyleOf(sheet, active.block)?.id ?? ''
  const fontFamilies = useFontFamilies(active.fontFamily)
  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label={t('document.styleAndFont.group')}>
      <ToolbarSelect
        label={t('document.styleAndFont.style')}
        value={currentStyle}
        options={withCurrent(styles, currentStyle)}
        onChange={(value) => chain().applyParagraphStyle(value).run()}
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

      <ToolbarButton icon="styles" label={t('document.styleAndFont.documentStyles')} onClick={onOpenStyles} />
    </ToolbarGroup>
  )
}
