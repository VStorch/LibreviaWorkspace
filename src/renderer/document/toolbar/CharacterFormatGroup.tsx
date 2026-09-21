import { useEditorState, type Editor } from '@tiptap/react'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ColorControl, ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { focusChain } from './focus-chain.js'

/** O que é propriedade do trecho de texto: as marcas e as cores. */
export function CharacterFormatGroup({ editor }: { readonly editor: Editor }): React.JSX.Element {
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      bold: current.isActive('bold'),
      italic: current.isActive('italic'),
      underline: current.isActive('underline'),
      strike: current.isActive('strike'),
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
    <ToolbarGroup label="Formatação do texto">
      <ToolbarButton
        icon="bold"
        label="Negrito"
        shortcut={shortcutHintOf(SHORTCUTS.bold)}
        active={active.bold}
        onClick={() => chain().toggleBold().run()}
      />
      <ToolbarButton
        icon="italic"
        label="Itálico"
        shortcut={shortcutHintOf(SHORTCUTS.italic)}
        active={active.italic}
        onClick={() => chain().toggleItalic().run()}
      />
      <ToolbarButton
        icon="underline"
        label="Sublinhado"
        shortcut={shortcutHintOf(SHORTCUTS.underline)}
        active={active.underline}
        onClick={() => chain().toggleUnderline().run()}
      />
      <ToolbarButton
        icon="strike"
        label="Tachado"
        active={active.strike}
        onClick={() => chain().toggleStrike().run()}
      />

      {/* Os atalhos anunciados são os do Word. Os padrões do Tiptap — `Ctrl+.`
          e `Ctrl+,` — continuam valendo, para o teclado em que o `=` não é
          uma tecla só. */}
      <ToolbarButton
        icon="superscript"
        label="Sobrescrito"
        shortcut={shortcutHintOf(SHORTCUTS.superscript)}
        active={active.superscript}
        onClick={() => chain().toggleSuperscript().run()}
      />
      <ToolbarButton
        icon="subscript"
        label="Subscrito"
        shortcut={shortcutHintOf(SHORTCUTS.subscript)}
        active={active.subscript}
        onClick={() => chain().toggleSubscript().run()}
      />
      <ToolbarButton
        icon="caps"
        label="Caixa alta"
        active={active.caps}
        onClick={() => chain().toggleCaps().run()}
      />
      <ToolbarButton
        icon="small-caps"
        label="Versalete"
        active={active.smallCaps}
        onClick={() => chain().toggleSmallCaps().run()}
      />

      <ColorControl
        icon="text-color"
        label="Cor do texto"
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
        label="Cor de fundo do texto"
        value={active.background}
        onChange={(value) => chain().setBackgroundColor(value).run()}
        onClear={() => chain().unsetBackgroundColor().run()}
      />
      <ColorControl
        icon="fill-color"
        label="Destaque"
        value={active.highlight}
        onChange={(value) => chain().setHighlight({ color: value }).run()}
        onClear={() => chain().unsetHighlight().run()}
      />
    </ToolbarGroup>
  )
}
