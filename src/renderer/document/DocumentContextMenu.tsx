import { Fragment } from 'react'
import type { IpcResult } from '@shared/ipc.js'
import { TABLE_ACTIONS, type TableAction } from '@shared/table-actions.js'
import { DictionaryScope, EditCommand, type ContextMenuTarget } from '@shared/types.js'
import { ContextMenu, ContextMenuItem, ContextMenuSeparator } from '../components/ContextMenu.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { NoteKind } from '@services/document/notes.js'

/** O que o botão direito oferece sobre uma lista. */
export type ListAction = 'restart' | 'continue' | 'setStart' | 'format'

export interface DocumentContextMenuProps {
  readonly target: ContextMenuTarget
  /** Só então as ações dela aparecem. */
  readonly inTable: boolean
  readonly onTableAction: (action: TableAction) => void
  readonly inList: 'bulletList' | 'orderedList' | null
  readonly onListAction: (action: ListAction) => void
  readonly onClose: () => void
  readonly onPasteWithoutFormat: () => void
  readonly onNewComment: () => void
  /** `null` fora de alteração. */
  readonly onRevision: ((accept: boolean) => void) | null
  /** `null` longe de nota (`noteAtCursor`). */
  readonly noteKind: NoteKind | null
  readonly onConvertNote: () => void
}

/**
 * Na ordem do Word: as sugestões do corretor primeiro, depois a área de
 * transferência. Os dados e as ações são do main, onde estão o corretor e o
 * `webContents`; colar sem formatação é a exceção e fica no editor.
 */
export function DocumentContextMenu(props: DocumentContextMenuProps): React.JSX.Element {
  const { target, inTable, inList, onRevision, noteKind, onClose } = props
  const readOnly = useWorkspace((state) => state.readOnly)
  const t = useT()
  const act = useMenuAction(onClose)

  return (
    <ContextMenu position={target} label={t('document.contextMenu.label')} onClose={onClose}>
      {target.misspelledWord !== '' && <SpellingItems target={target} act={act} onClose={onClose} />}

      <ClipboardItems {...props} act={act} readOnly={readOnly} />

      {onRevision !== null && !readOnly && (
        <>
          <ContextMenuSeparator />
          <ContextMenuItem onClick={closing(onClose, () => onRevision(true))}>
            {t('revisions.accept')}
          </ContextMenuItem>
          <ContextMenuItem onClick={closing(onClose, () => onRevision(false))}>
            {t('revisions.reject')}
          </ContextMenuItem>
        </>
      )}

      {noteKind !== null && !readOnly && (
        <>
          <ContextMenuSeparator />
          <ContextMenuItem onClick={closing(onClose, props.onConvertNote)}>
            {t(
              noteKind === NoteKind.Footnote
                ? 'document.contextMenu.toEndnote'
                : 'document.contextMenu.toFootnote',
            )}
          </ContextMenuItem>
        </>
      )}

      {inList !== null && !readOnly && (
        <ListItems inList={inList} onListAction={props.onListAction} onClose={onClose} />
      )}

      {/* "Inserir tabela" fica de fora: tabela dentro de tabela mora no menu "Tabela". */}
      {inTable && !readOnly && <TableItems onTableAction={props.onTableAction} onClose={onClose} />}
    </ContextMenu>
  )
}

type MenuAction = (run: () => Promise<IpcResult<unknown>>) => () => void

/** Toda ação fecha o menu, inclusive quando falha, para o erro ficar visível. */
function useMenuAction(onClose: () => void): MenuAction {
  const showError = useWorkspace((state) => state.showError)
  return (run) => () => {
    onClose()
    void run().then((result) => {
      if (!result.ok) showError(result.error)
    })
  }
}

const closing = (onClose: () => void, run: () => void) => (): void => {
  onClose()
  run()
}

function SpellingItems({
  target,
  act,
  onClose,
}: {
  target: ContextMenuTarget
  act: MenuAction
  onClose: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      {target.dictionarySuggestions.length === 0 ? (
        // Um item apagado, e não nenhum: senão pareceria que o menu quebrou.
        <ContextMenuItem disabled onClick={onClose}>
          {t('document.contextMenu.noSuggestions')}
        </ContextMenuItem>
      ) : (
        target.dictionarySuggestions.map((suggestion) => (
          <ContextMenuItem
            key={suggestion}
            strong
            onClick={act(() => window.api.spell.replace({ word: suggestion }))}
          >
            {suggestion}
          </ContextMenuItem>
        ))
      )}

      <ContextMenuSeparator />

      <ContextMenuItem
        onClick={act(() =>
          window.api.spell.addWord({ word: target.misspelledWord, scope: DictionaryScope.Permanent }),
        )}
      >
        {t('document.contextMenu.addToDictionary')}
      </ContextMenuItem>
      {/* "Ignorar" vale até fechar o aplicativo: o Chromium não tem lista de ignorados. */}
      <ContextMenuItem
        onClick={act(() =>
          window.api.spell.addWord({ word: target.misspelledWord, scope: DictionaryScope.Session }),
        )}
      >
        {t('document.contextMenu.ignoreSession')}
      </ContextMenuItem>

      <ContextMenuSeparator />
    </>
  )
}

function ClipboardItems({
  target,
  act,
  readOnly,
  onClose,
  onPasteWithoutFormat,
  onNewComment,
}: DocumentContextMenuProps & { act: MenuAction; readOnly: boolean }): React.JSX.Element {
  const t = useT()
  return (
    <>
      <ContextMenuItem
        disabled={!target.canCut || readOnly}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Cut }))}
      >
        {t('menu.edit.cut')}
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canCopy}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Copy }))}
      >
        {t('menu.edit.copy')}
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canPaste || readOnly}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Paste }))}
      >
        {t('menu.edit.paste')}
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canPaste || readOnly}
        onClick={closing(onClose, onPasteWithoutFormat)}
      >
        {t('menu.edit.pasteWithoutFormat')}
      </ContextMenuItem>

      {/* Logo depois da área de transferência, como o "Novo comentário" do Word. */}
      <ContextMenuSeparator />
      <ContextMenuItem disabled={readOnly} onClick={closing(onClose, onNewComment)}>
        {t('comments.new')}
      </ContextMenuItem>
    </>
  )
}

/** Reiniciar e continuar só fazem sentido em lista numerada. */
function ListItems({
  inList,
  onListAction,
  onClose,
}: {
  inList: 'bulletList' | 'orderedList'
  onListAction: (action: ListAction) => void
  onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const item = (action: ListAction, label: string): React.JSX.Element => (
    <ContextMenuItem onClick={closing(onClose, () => onListAction(action))}>{label}</ContextMenuItem>
  )
  return (
    <>
      <ContextMenuSeparator />
      {inList === 'orderedList' && (
        <>
          {item('restart', t('document.lists.restart'))}
          {item('continue', t('document.lists.continue'))}
          {item('setStart', t('document.lists.setStart'))}
        </>
      )}
      {item('format', t('document.lists.format'))}
    </>
  )
}

function TableItems({
  onTableAction,
  onClose,
}: {
  onTableAction: (action: TableAction) => void
  onClose: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      {TABLE_ACTIONS.filter((action) => action.needsTable).map((action, index, list) => (
        <Fragment key={action.id}>
          {(index === 0 || list[index - 1]?.group !== action.group) && <ContextMenuSeparator />}
          <ContextMenuItem onClick={closing(onClose, () => onTableAction(action.id))}>
            {t(action.labelKey)}
          </ContextMenuItem>
        </Fragment>
      ))}
    </>
  )
}
