import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { firstFamilyOf } from '@services/document/font-list.js'
import {
  ColorControl,
  ToolbarButton,
  ToolbarGroup,
  ToolbarSelect,
  ToolbarSeparator,
} from '../../components/ToolbarControls.js'
import { setPreference, usePreferences } from '../../state/preferences.js'
import { useWorkspace } from '../../state/workspace.js'
import { blockLineHeightOf } from '../extensions/paragraph-commands.js'
import { LinkDialog } from './LinkDialog.js'
import { ParagraphDialog } from './ParagraphDialog.js'
import { BLOCK_STYLES, FONT_SIZES, LINE_HEIGHTS, withCurrent } from './toolbar-options.js'
import { useFontFamilies } from './useFontFamilies.js'

interface DocumentToolbarProps {
  readonly editor: Editor
  readonly onOpenFind: () => void
  readonly onOpenPageSetup: () => void
  /** Aberto de fora também: o menu nativo tem "Formatar → Parágrafo…". */
  readonly paragraphOpen: boolean
  readonly onParagraphOpenChange: (open: boolean) => void
}

export function DocumentToolbar({
  editor,
  onOpenFind,
  onOpenPageSetup,
  paragraphOpen,
  onParagraphOpenChange,
}: DocumentToolbarProps): React.JSX.Element {
  const [linkDialogOpen, setLinkDialogOpen] = useState(false)
  const showError = useWorkspace((state) => state.showError)
  const printPreview = useWorkspace((state) => state.printPreview)
  // Do main, que é o dono da preferência: clicar aqui e clicar no item do menu
  // "Exibir" mudam a mesma chave, e os dois ficam marcados juntos.
  const invisibleCharacters = usePreferences((state) => state.preferences.invisibleCharacters)

  // Reavalia só o que a barra desenha, a cada transação do editor.
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
      bulletList: current.isActive('bulletList'),
      orderedList: current.isActive('orderedList'),
      alignLeft: current.isActive({ textAlign: 'left' }),
      alignCenter: current.isActive({ textAlign: 'center' }),
      alignRight: current.isActive({ textAlign: 'right' }),
      alignJustify: current.isActive({ textAlign: 'justify' }),
      link: current.isActive('link'),
      inTable: current.isActive('table'),
      heading: current.isActive('heading') ? String(current.getAttributes('heading')['level'] ?? '') : '',
      // Só o nome da fonte: o que vem do documento é uma pilha de CSS, com a
      // substituta genérica atrás, e é o nome que a lista aqui conhece.
      fontFamily: firstFamilyOf(String(current.getAttributes('textStyle')['fontFamily'] ?? '')),
      fontSize: String(current.getAttributes('textStyle')['fontSize'] ?? '').replace('pt', ''),
      // Do **bloco**, e não da marca de texto: no OOXML a entrelinha é
      // propriedade do parágrafo, e não existe `w:line` dentro de um `w:rPr`.
      // Enquanto este seletor escrevia na marca, escolher "Duplo" aqui era perda
      // garantida — o gravador não tinha onde pôr a medida e a anotava no
      // inventário.
      lineHeight: blockLineHeightOf(current),
      color: String(current.getAttributes('textStyle')['color'] ?? '#000000'),
      background: String(current.getAttributes('textStyle')['backgroundColor'] ?? '#ffff00'),
      highlight: String(current.getAttributes('highlight')['color'] ?? '#ffff00'),
    }),
  })

  const fontFamilies = useFontFamilies(active.fontFamily)
  const chain = () => editor.chain().focus()

  function applyBlockStyle(value: string): void {
    if (value === 'paragraph') chain().setParagraph().run()
    else
      chain()
        .toggleHeading({ level: Number(value) as 1 | 2 | 3 | 4 })
        .run()
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
    <div className="toolbar" role="toolbar" aria-label="Formatação do documento">
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

      <ToolbarSeparator />

      <ToolbarGroup label="Formatação do texto">
        <ToolbarButton
          icon="bold"
          label="Negrito"
          shortcut="Ctrl+B"
          active={active.bold}
          onClick={() => chain().toggleBold().run()}
        />
        <ToolbarButton
          icon="italic"
          label="Itálico"
          shortcut="Ctrl+I"
          active={active.italic}
          onClick={() => chain().toggleItalic().run()}
        />
        <ToolbarButton
          icon="underline"
          label="Sublinhado"
          shortcut="Ctrl+U"
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
          shortcut="Ctrl+Shift+="
          active={active.superscript}
          onClick={() => chain().toggleSuperscript().run()}
        />
        <ToolbarButton
          icon="subscript"
          label="Subscrito"
          shortcut="Ctrl+="
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

      <ToolbarSeparator />

      <ToolbarGroup label="Parágrafo">
        <ToolbarButton
          icon="align-left"
          label="Alinhar à esquerda"
          shortcut="Ctrl+L"
          active={active.alignLeft}
          onClick={() => chain().setTextAlign('left').run()}
        />
        <ToolbarButton
          icon="align-center"
          label="Centralizar"
          shortcut="Ctrl+E"
          active={active.alignCenter}
          onClick={() => chain().setTextAlign('center').run()}
        />
        <ToolbarButton
          icon="align-right"
          label="Alinhar à direita"
          shortcut="Ctrl+R"
          active={active.alignRight}
          onClick={() => chain().setTextAlign('right').run()}
        />
        <ToolbarButton
          icon="align-justify"
          label="Justificar"
          shortcut="Ctrl+J"
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

      <ToolbarSeparator />

      <ToolbarGroup label="Listas e recuo">
        <ToolbarButton
          icon="bullet-list"
          label="Lista com marcadores"
          active={active.bulletList}
          onClick={() => chain().toggleBulletList().run()}
        />
        <ToolbarButton
          icon="ordered-list"
          label="Lista numerada"
          active={active.orderedList}
          onClick={() => chain().toggleOrderedList().run()}
        />
        <ToolbarButton
          icon="outdent"
          label="Diminuir recuo"
          shortcut="Ctrl+["
          onClick={() => chain().outdent().run()}
        />
        <ToolbarButton
          icon="indent"
          label="Aumentar recuo"
          shortcut="Ctrl+]"
          onClick={() => chain().indent().run()}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label="Inserir">
        <ToolbarButton
          icon="table"
          label={active.inTable ? 'Remover tabela' : 'Inserir tabela'}
          active={active.inTable}
          onClick={() =>
            active.inTable
              ? chain().deleteTable().run()
              : chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
          }
        />
        <ToolbarButton icon="image" label="Inserir imagem" onClick={() => void insertImage()} />
        <ToolbarButton
          icon="link"
          label="Inserir link"
          active={active.link}
          onClick={() => setLinkDialogOpen(true)}
        />
        <ToolbarButton
          icon="page-break"
          label="Quebra de página"
          shortcut="Ctrl+Enter"
          onClick={() => chain().setPageBreak().run()}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label="Página">
        <ToolbarButton
          icon="formatting-marks"
          label="Marcas de formatação"
          shortcut="Ctrl+F10"
          active={invisibleCharacters}
          onClick={() => void setPreference({ invisibleCharacters: !invisibleCharacters })}
        />
        <ToolbarButton icon="search" label="Localizar e substituir" shortcut="Ctrl+F" onClick={onOpenFind} />
        <ToolbarButton icon="page-setup" label="Configuração de página" onClick={onOpenPageSetup} />
        {/* Como o editor não pagina ao vivo (§6.3 do plano), a prévia é o que
          responde "onde as páginas quebram" — e por isso fica à mão. */}
        <ToolbarButton
          icon="print-preview"
          label="Visualizar impressão"
          onClick={() => void printPreview()}
        />
      </ToolbarGroup>

      {linkDialogOpen && <LinkDialog editor={editor} onClose={() => setLinkDialogOpen(false)} />}
      {paragraphOpen && <ParagraphDialog editor={editor} onClose={() => onParagraphOpenChange(false)} />}
    </div>
  )
}
