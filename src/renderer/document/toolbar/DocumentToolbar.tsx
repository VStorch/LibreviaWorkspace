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
  readonly onOpenStyles: () => void
  /** The native menu also opens the paragraph dialog. */
  readonly paragraphOpen: boolean
  readonly onParagraphOpenChange: (open: boolean) => void
  /** The native and context menus also open them: the state lives above. */
  readonly onOpenTable: () => void
  readonly onOpenImageProperties: () => void
  readonly onOpenListFormat: () => void
}

/** Only the group order and the dialogs; each group watches only what it draws. */
export function DocumentToolbar({
  editor,
  onOpenFind,
  onOpenPageSetup,
  onOpenStyles,
  paragraphOpen,
  onParagraphOpenChange,
  onOpenTable,
  onOpenImageProperties,
  onOpenListFormat,
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

      <ListAndIndentGroup editor={editor} onOpenListFormat={onOpenListFormat} />

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
