import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { ToolbarSeparator } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { CharacterFormatGroup } from './CharacterFormatGroup.js'
import { InsertGroup } from './InsertGroup.js'
import { LinkDialog } from './LinkDialog.js'
import { ListAndIndentGroup } from './ListAndIndentGroup.js'
import { PageGroup } from './PageGroup.js'
import { ParagraphDialog } from './ParagraphDialog.js'
import { ParagraphGroup } from './ParagraphGroup.js'
import { StyleAndFontGroup } from './StyleAndFontGroup.js'

interface DocumentToolbarProps {
  readonly editor: Editor
  readonly onOpenFind: () => void
  readonly onOpenPageSetup: () => void
  /** O painel de estilos do documento, só de leitura nesta fase. */
  readonly onOpenStyles: () => void
  /** Aberto de fora também: o menu nativo tem "Formatar → Parágrafo…". */
  readonly paragraphOpen: boolean
  readonly onParagraphOpenChange: (open: boolean) => void
  /**
   * Os diálogos de tabela e de imagem são desenhados pelo editor, e não aqui: o
   * menu nativo e o menu de contexto também os abrem, então o estado deles mora
   * um nível acima.
   */
  readonly onOpenTable: () => void
  readonly onOpenImageProperties: () => void
}

/**
 * A barra de ferramentas do documento.
 *
 * Aqui mora só a ordem dos grupos e os diálogos que eles abrem. Cada grupo é um
 * arquivo, e cada arquivo observa no editor apenas o que os seus próprios botões
 * desenham — a barra inteira reavaliada a cada tecla digitada era o que fazia um
 * recurso novo custar linhas neste arquivo em vez de custar um arquivo novo.
 *
 * Os diálogos ficam fora dos grupos, no fim da barra, porque são um só por vez e
 * se posicionam em relação à janela.
 */
export function DocumentToolbar({
  editor,
  onOpenFind,
  onOpenPageSetup,
  onOpenStyles,
  paragraphOpen,
  onParagraphOpenChange,
  onOpenTable,
  onOpenImageProperties,
}: DocumentToolbarProps): React.JSX.Element {
  const t = useT()
  const [linkDialogOpen, setLinkDialogOpen] = useState(false)

  return (
    <div className="toolbar" role="toolbar" aria-label={t('document.toolbar.label')}>
      <StyleAndFontGroup editor={editor} onOpenStyles={onOpenStyles} />

      <ToolbarSeparator />

      <CharacterFormatGroup editor={editor} />

      <ToolbarSeparator />

      <ParagraphGroup
        editor={editor}
        paragraphOpen={paragraphOpen}
        onParagraphOpenChange={onParagraphOpenChange}
      />

      <ToolbarSeparator />

      <ListAndIndentGroup editor={editor} />

      <ToolbarSeparator />

      <InsertGroup
        editor={editor}
        onOpenLink={() => setLinkDialogOpen(true)}
        onOpenTable={onOpenTable}
        onOpenImageProperties={onOpenImageProperties}
      />

      <ToolbarSeparator />

      <PageGroup onOpenFind={onOpenFind} onOpenPageSetup={onOpenPageSetup} />

      {linkDialogOpen && <LinkDialog editor={editor} onClose={() => setLinkDialogOpen(false)} />}
      {paragraphOpen && <ParagraphDialog editor={editor} onClose={() => onParagraphOpenChange(false)} />}
    </div>
  )
}
