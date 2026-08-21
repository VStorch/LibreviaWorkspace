import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import {
  ColorControl,
  ToolbarButton,
  ToolbarGroup,
  ToolbarSelect,
  ToolbarSeparator,
} from '../../components/ToolbarControls.js'
import { useWorkspace } from '../../state/workspace.js'
import { LinkDialog } from './LinkDialog.js'

/** Famílias disponíveis. A Fase 3 embute fontes métricas do Word (§6.4). */
const FONT_FAMILIES = [
  { value: '', label: 'Fonte padrão' },
  { value: 'Arial, Liberation Sans, sans-serif', label: 'Arial' },
  { value: 'Calibri, Carlito, sans-serif', label: 'Calibri' },
  { value: 'Cambria, Caladea, serif', label: 'Cambria' },
  { value: 'Georgia, serif', label: 'Georgia' },
  { value: 'Times New Roman, Liberation Serif, serif', label: 'Times New Roman' },
  { value: 'Verdana, sans-serif', label: 'Verdana' },
  { value: 'Courier New, Liberation Mono, monospace', label: 'Courier New' },
] as const

const FONT_SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32', '36', '48', '72']

const LINE_HEIGHTS = [
  { value: '', label: 'Simples' },
  { value: '1.15', label: '1,15' },
  { value: '1.5', label: '1,5' },
  { value: '2', label: 'Duplo' },
] as const

const BLOCK_STYLES = [
  { value: 'paragraph', label: 'Texto normal' },
  { value: '1', label: 'Título 1' },
  { value: '2', label: 'Título 2' },
  { value: '3', label: 'Título 3' },
  { value: '4', label: 'Título 4' },
] as const

interface DocumentToolbarProps {
  readonly editor: Editor
  readonly onOpenFind: () => void
  readonly onOpenPageSetup: () => void
}

export function DocumentToolbar({
  editor,
  onOpenFind,
  onOpenPageSetup,
}: DocumentToolbarProps): React.JSX.Element {
  const [linkDialogOpen, setLinkDialogOpen] = useState(false)
  const showError = useWorkspace((state) => state.showError)
  const printPreview = useWorkspace((state) => state.printPreview)

  // Reavalia só o que a barra desenha, a cada transação do editor.
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => ({
      bold: current.isActive('bold'),
      italic: current.isActive('italic'),
      underline: current.isActive('underline'),
      strike: current.isActive('strike'),
      bulletList: current.isActive('bulletList'),
      orderedList: current.isActive('orderedList'),
      alignLeft: current.isActive({ textAlign: 'left' }),
      alignCenter: current.isActive({ textAlign: 'center' }),
      alignRight: current.isActive({ textAlign: 'right' }),
      alignJustify: current.isActive({ textAlign: 'justify' }),
      link: current.isActive('link'),
      inTable: current.isActive('table'),
      heading: current.isActive('heading') ? String(current.getAttributes('heading')['level'] ?? '') : '',
      fontFamily: String(current.getAttributes('textStyle')['fontFamily'] ?? ''),
      fontSize: String(current.getAttributes('textStyle')['fontSize'] ?? '').replace('pt', ''),
      lineHeight: String(current.getAttributes('textStyle')['lineHeight'] ?? ''),
      color: String(current.getAttributes('textStyle')['color'] ?? '#000000'),
      highlight: String(current.getAttributes('highlight')['color'] ?? '#ffff00'),
    }),
  })

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
          options={FONT_FAMILIES}
          onChange={(value) =>
            value === '' ? chain().unsetFontFamily().run() : chain().setFontFamily(value).run()
          }
          width={150}
        />

        <ToolbarSelect
          label="Tamanho"
          value={active.fontSize}
          options={[{ value: '', label: '—' }, ...FONT_SIZES.map((size) => ({ value: size, label: size }))]}
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

        <ColorControl
          icon="text-color"
          label="Cor do texto"
          value={active.color}
          onChange={(value) => chain().setColor(value).run()}
          onClear={() => chain().unsetColor().run()}
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
          active={active.alignLeft}
          onClick={() => chain().setTextAlign('left').run()}
        />
        <ToolbarButton
          icon="align-center"
          label="Centralizar"
          active={active.alignCenter}
          onClick={() => chain().setTextAlign('center').run()}
        />
        <ToolbarButton
          icon="align-right"
          label="Alinhar à direita"
          active={active.alignRight}
          onClick={() => chain().setTextAlign('right').run()}
        />
        <ToolbarButton
          icon="align-justify"
          label="Justificar"
          active={active.alignJustify}
          onClick={() => chain().setTextAlign('justify').run()}
        />

        <ToolbarSelect
          label="Espaçamento entre linhas"
          value={active.lineHeight}
          options={LINE_HEIGHTS}
          onChange={(value) =>
            value === '' ? chain().unsetLineHeight().run() : chain().setLineHeight(value).run()
          }
          width={100}
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
    </div>
  )
}
