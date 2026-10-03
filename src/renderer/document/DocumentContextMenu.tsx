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

/**
 * Na ordem do Word: as sugestões do corretor primeiro, depois a área de
 * transferência. Os dados e as ações são do main, onde estão o corretor e o
 * `webContents`; colar sem formatação é a exceção e fica no editor.
 */
export function DocumentContextMenu({
  target,
  inTable,
  onTableAction,
  inList,
  onListAction,
  onClose,
  onPasteWithoutFormat,
  onNewComment,
  onRevision,
  noteKind,
  onConvertNote,
}: {
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
}): React.JSX.Element {
  const showError = useWorkspace((state) => state.showError)
  const readOnly = useWorkspace((state) => state.readOnly)
  const t = useT()

  /** Toda ação fecha o menu, inclusive quando falha, para o erro ficar visível. */
  const act = (run: () => Promise<IpcResult<unknown>>) => () => {
    onClose()
    void run().then((result) => {
      if (!result.ok) showError(result.error)
    })
  }

  const misspelled = target.misspelledWord !== ''

  return (
    <ContextMenu position={target} label={t('document.contextMenu.label')} onClose={onClose}>
      {misspelled && (
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
              window.api.spell.addWord({
                word: target.misspelledWord,
                scope: DictionaryScope.Permanent,
              }),
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
      )}

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
        onClick={() => {
          onClose()
          onPasteWithoutFormat()
        }}
      >
        {t('menu.edit.pasteWithoutFormat')}
      </ContextMenuItem>

      {/* Logo depois da área de transferência, como o "Novo comentário" do Word. */}
      <ContextMenuSeparator />
      <ContextMenuItem
        disabled={readOnly}
        onClick={() => {
          onClose()
          onNewComment()
        }}
      >
        {t('comments.new')}
      </ContextMenuItem>

      {onRevision !== null && !readOnly && (
        <>
          <ContextMenuSeparator />
          <ContextMenuItem
            onClick={() => {
              onClose()
              onRevision(true)
            }}
          >
            {t('revisions.accept')}
          </ContextMenuItem>
          <ContextMenuItem
            onClick={() => {
              onClose()
              onRevision(false)
            }}
          >
            {t('revisions.reject')}
          </ContextMenuItem>
        </>
      )}

      {noteKind !== null && !readOnly && (
        <>
          <ContextMenuSeparator />
          <ContextMenuItem
            onClick={() => {
              onClose()
              onConvertNote()
            }}
          >
            {t(
              noteKind === NoteKind.Footnote
                ? 'document.contextMenu.toEndnote'
                : 'document.contextMenu.toFootnote',
            )}
          </ContextMenuItem>
        </>
      )}

      {/* Reiniciar e continuar só fazem sentido em lista numerada. */}
      {inList !== null && !readOnly && (
        <>
          <ContextMenuSeparator />
          {inList === 'orderedList' && (
            <>
              <ContextMenuItem
                onClick={() => {
                  onClose()
                  onListAction('restart')
                }}
              >
                {t('document.lists.restart')}
              </ContextMenuItem>
              <ContextMenuItem
                onClick={() => {
                  onClose()
                  onListAction('continue')
                }}
              >
                {t('document.lists.continue')}
              </ContextMenuItem>
              <ContextMenuItem
                onClick={() => {
                  onClose()
                  onListAction('setStart')
                }}
              >
                {t('document.lists.setStart')}
              </ContextMenuItem>
            </>
          )}
          <ContextMenuItem
            onClick={() => {
              onClose()
              onListAction('format')
            }}
          >
            {t('document.lists.format')}
          </ContextMenuItem>
        </>
      )}

      {/* "Inserir tabela" fica de fora: tabela dentro de tabela mora no menu "Tabela". */}
      {inTable && !readOnly && (
        <>
          {TABLE_ACTIONS.filter((action) => action.needsTable).map((action, index, list) => (
            <Fragment key={action.id}>
              {(index === 0 || list[index - 1]?.group !== action.group) && <ContextMenuSeparator />}
              <ContextMenuItem
                onClick={() => {
                  onClose()
                  onTableAction(action.id)
                }}
              >
                {t(action.labelKey)}
              </ContextMenuItem>
            </Fragment>
          ))}
        </>
      )}
    </ContextMenu>
  )
}
