import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ColorControl, ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { markVisiblyOn } from '../extensions/style-commands.js'
import { focusChain } from './focus-chain.js'

/** O que é propriedade do trecho de texto: as marcas e as cores. */
export function CharacterFormatGroup({ editor }: { readonly editor: Editor }): React.JSX.Element {
  const t = useT()
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      // O que aparece, com o estilo por baixo: o título negrito sem marca está
      // negrito, e o trecho com a marca "desligado" não está.
      bold: markVisiblyOn(current.state.tr, current.storage.paragraphCommands.styles, 'bold'),
      italic: markVisiblyOn(current.state.tr, current.storage.paragraphCommands.styles, 'italic'),
      underline: markVisiblyOn(current.state.tr, current.storage.paragraphCommands.styles, 'underline'),
      strike: markVisiblyOn(current.state.tr, current.storage.paragraphCommands.styles, 'strike'),
      superscript: current.isActive('superscript'),
      subscript: current.isActive('subscript'),
      caps: current.isActive('caps'),
      smallCaps: current.isActive('smallCaps'),
      color: String(current.getAttributes('textStyle')['color'] ?? '#000000'),
      background: String(current.getAttributes('textStyle')['backgroundColor'] ?? '#ffff00'),
      highlight: String(current.getAttributes('highlight')['color'] ?? '#ffff00'),
    }),
  })

  const chain = () => focusChain(editor)

  return (
    <ToolbarGroup label={t('document.characterFormat.group')}>
      <ToolbarButton
        icon="bold"
        label={t('document.characterFormat.bold')}
        shortcut={shortcutHintOf(SHORTCUTS.bold)}
        active={active.bold}
        onClick={() => chain().toggleInheritedMark('bold').run()}
      />
      <ToolbarButton
        icon="italic"
        label={t('document.characterFormat.italic')}
        shortcut={shortcutHintOf(SHORTCUTS.italic)}
        active={active.italic}
        onClick={() => chain().toggleInheritedMark('italic').run()}
      />
      <ToolbarButton
        icon="underline"
        label={t('document.characterFormat.underline')}
        shortcut={shortcutHintOf(SHORTCUTS.underline)}
        active={active.underline}
        onClick={() => chain().toggleInheritedMark('underline').run()}
      />
      <ToolbarButton
        icon="strike"
        label={t('document.characterFormat.strikethrough')}
        active={active.strike}
        onClick={() => chain().toggleInheritedMark('strike').run()}
      />

      {/* Os atalhos anunciados são os do Word. Os padrões do Tiptap — `Ctrl+.`
          e `Ctrl+,` — continuam valendo, para o teclado em que o `=` não é
          uma tecla só. */}
      <ToolbarButton
        icon="superscript"
        label={t('document.characterFormat.superscript')}
        shortcut={shortcutHintOf(SHORTCUTS.superscript)}
        active={active.superscript}
        onClick={() => chain().toggleSuperscript().run()}
      />
      <ToolbarButton
        icon="subscript"
        label={t('document.characterFormat.subscript')}
        shortcut={shortcutHintOf(SHORTCUTS.subscript)}
        active={active.subscript}
        onClick={() => chain().toggleSubscript().run()}
      />
      <ToolbarButton
        icon="caps"
        label={t('document.characterFormat.uppercase')}
        active={active.caps}
        onClick={() => chain().toggleCaps().run()}
      />
      <ToolbarButton
        icon="small-caps"
        label={t('document.characterFormat.smallCaps')}
        active={active.smallCaps}
        onClick={() => chain().toggleSmallCaps().run()}
      />

      <ColorControl
        icon="text-color"
        label={t('document.characterFormat.textColor')}
        value={active.color}
        onChange={(value) => chain().setColor(value).run()}
        onClear={() => chain().unsetColor().run()}
      />
      {/* Duas cores de fundo, e não uma por engano: "Destaque" é o marca-texto
          do Word (`w:highlight`, catorze cores fixas) e esta é o sombreamento
          do trecho (`w:shd`, cor livre). No arquivo são propriedades
          diferentes, e um documento importado pode trazer as duas. */}
      <ColorControl
        icon="text-background"
        label={t('document.characterFormat.backgroundColor')}
        value={active.background}
        onChange={(value) => chain().setBackgroundColor(value).run()}
        onClear={() => chain().unsetBackgroundColor().run()}
      />
      <ColorControl
        icon="fill-color"
        label={t('document.characterFormat.highlight')}
        value={active.highlight}
        onChange={(value) => chain().setHighlight({ color: value }).run()}
        onClear={() => chain().unsetHighlight().run()}
      />
    </ToolbarGroup>
  )
}
