import { Fragment } from 'react'
import type { Editor } from '@tiptap/react'
import type { ContextMenuTarget } from '@shared/types.js'
import type { ResolvedSections } from '@services/document/sections.js'
import { TableDialog } from './toolbar/TableDialog.js'
import { TablePropertiesDialog } from './toolbar/TablePropertiesDialog.js'
import { ImageDialog } from './toolbar/ImageDialog.js'
import { DocumentContextMenu } from './DocumentContextMenu.js'
import { ListFormatDialog, ListStartDialog } from './ListFormatDialog.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { SpecialCharsDialog } from './SpecialCharsDialog.js'
import { MathDialog } from './MathDialog.js'
import { StylesPanel } from './StylesPanel.js'
import { BookmarkDialog } from './BookmarkDialog.js'
import { ColumnsDialog } from './ColumnsDialog.js'
import { CaptionDialog } from './CaptionDialog.js'
import { CrossReferenceDialog } from './CrossReferenceDialog.js'
import { WordCountDialog } from './WordCountDialog.js'
import { PropertiesDialog } from './PropertiesDialog.js'
import { AuthorNameDialog } from './AuthorNameDialog.js'
import { insertComment } from './comment-commands.js'
import { hasChangeAtCursor, settleChange } from './revision-commands.js'
import { convertNote, noteAtCursor } from './note-commands.js'
import { sectionAtCursor } from './section-commands.js'
import type { EquationTarget } from './math-commands.js'
import type { ReferenceContext } from './references.js'
import type { SearchStatus } from './extensions/search-replace.js'
import type { EditorCommand } from './editor-commands.js'
import type { EditorDialogs as OpenDialogs } from './useEditorCommands.js'

type DialogKey = keyof OpenDialogs

export interface DialogContext {
  readonly editor: Editor
  readonly close: (dialog: DialogKey) => () => void
  readonly searchStatus: SearchStatus
  readonly resolved: ResolvedSections
  readonly equationTarget: EquationTarget
  readonly readOnly: boolean
  readonly referenceContext: () => ReferenceContext
}

type SimpleDialog = (props: { editor: Editor; onClose: () => void }) => React.JSX.Element | null

const simple =
  (key: DialogKey, Dialog: SimpleDialog) =>
  (context: DialogContext): React.JSX.Element => (
    <Dialog editor={context.editor} onClose={context.close(key)} />
  )

/** In the order they enter the tree. */
const DIALOGS: readonly { key: DialogKey; render: (context: DialogContext) => React.JSX.Element }[] = [
  {
    key: 'find',
    render: (c) => <FindReplacePanel editor={c.editor} status={c.searchStatus} onClose={c.close('find')} />,
  },
  {
    key: 'pageSetup',
    render: (c) => (
      <PageSetupPanel
        onClose={c.close('pageSetup')}
        resolved={c.resolved}
        sectionIndex={sectionAtCursor(c.editor, c.resolved)}
      />
    ),
  },
  { key: 'wordCount', render: simple('wordCount', WordCountDialog) },
  { key: 'properties', render: simple('properties', PropertiesDialog) },
  { key: 'authorName', render: simple('authorName', AuthorNameDialog) },
  { key: 'styles', render: simple('styles', StylesPanel) },
  { key: 'specialCharacter', render: simple('specialCharacter', SpecialCharsDialog) },
  {
    key: 'equation',
    render: (c) => (
      <MathDialog
        editor={c.editor}
        target={c.equationTarget}
        readOnly={c.readOnly}
        onClose={c.close('equation')}
      />
    ),
  },
  { key: 'table', render: simple('table', TableDialog) },
  { key: 'tableProperties', render: simple('tableProperties', TablePropertiesDialog) },
  { key: 'imageProperties', render: simple('imageProperties', ImageDialog) },
  { key: 'listFormat', render: simple('listFormat', ListFormatDialog) },
  { key: 'listStart', render: simple('listStart', ListStartDialog) },
  { key: 'columns', render: simple('columns', ColumnsDialog) },
  { key: 'bookmark', render: simple('bookmark', BookmarkDialog) },
  {
    key: 'caption',
    render: (c) => (
      <CaptionDialog editor={c.editor} context={c.referenceContext} onClose={c.close('caption')} />
    ),
  },
  {
    key: 'crossReference',
    render: (c) => (
      <CrossReferenceDialog
        editor={c.editor}
        context={c.referenceContext}
        onClose={c.close('crossReference')}
      />
    ),
  },
]

export function EditorDialogs({
  open,
  context,
}: {
  open: OpenDialogs
  context: DialogContext
}): React.JSX.Element {
  return (
    <>
      {DIALOGS.filter(({ key }) => open[key]).map(({ key, render }) => (
        <Fragment key={key}>{render(context)}</Fragment>
      ))}
    </>
  )
}

export interface EditorContextMenuProps {
  readonly editor: Editor
  readonly target: ContextMenuTarget
  readonly run: (command: EditorCommand) => void
  readonly setDialog: (dialog: DialogKey, open: boolean) => void
  readonly pasteWithoutFormat: () => Promise<void>
  readonly onClose: () => void
}

export function EditorContextMenu({
  editor,
  target,
  run,
  setDialog,
  pasteWithoutFormat,
  onClose,
}: EditorContextMenuProps): React.JSX.Element {
  const contextNote = noteAtCursor(editor)
  return (
    <DocumentContextMenu
      target={target}
      // Outside a table, "merge cells" has nothing to merge.
      inTable={editor.isActive('table')}
      onTableAction={run}
      // As in Word, by right-clicking the item to restart.
      inList={
        editor.isActive('orderedList') ? 'orderedList' : editor.isActive('bulletList') ? 'bulletList' : null
      }
      onListAction={(action) => {
        if (action === 'restart') editor.chain().focus().restartListNumbering(1).run()
        else if (action === 'continue') editor.chain().focus().continueListNumbering().run()
        else setDialog(action === 'setStart' ? 'listStart' : 'listFormat', true)
      }}
      onClose={onClose}
      onPasteWithoutFormat={() => void pasteWithoutFormat()}
      onNewComment={() => insertComment(editor)}
      onRevision={hasChangeAtCursor(editor) ? (accept) => void settleChange(editor, accept) : null}
      noteKind={contextNote?.kind ?? null}
      onConvertNote={() => {
        if (contextNote !== null) convertNote(editor, contextNote.pos)
      }}
    />
  )
}
