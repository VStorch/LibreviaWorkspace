import { Fragment } from 'react'
import type { IpcResult } from '@shared/ipc.js'
import { TABLE_ACTIONS, type TableAction } from '@shared/table-actions.js'
import { DictionaryScope, EditCommand, type ContextMenuTarget } from '@shared/types.js'
import { ContextMenu, ContextMenuItem, ContextMenuSeparator } from '../components/ContextMenu.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Menu de contexto do editor de documentos.
 *
 * O que se espera do botão direito num editor de texto, na ordem do Word: as
 * sugestões do corretor primeiro — é por elas que se clica com o botão direito
 * numa palavra sublinhada — e depois a área de transferência.
 *
 * Os dados de onde se clicou vêm do processo main, porque só lá o Chromium conta
 * qual palavra ele marcou e o que sugere (ver `src/main/context-menu.ts`). As
 * ações voltam para lá pelo mesmo motivo: recortar, copiar, colar e trocar a
 * palavra errada são operações do `webContents`.
 *
 * Colar sem formatação é a exceção e fica no editor: o texto vem do main, mas
 * quem o transforma em parágrafos é o documento.
 */
/** O que o botão direito oferece sobre uma lista. */
export type ListAction = 'restart' | 'continue' | 'setStart' | 'format'

export function DocumentContextMenu({
  target,
  inTable,
  onTableAction,
  inList,
  onListAction,
  onClose,
  onPasteWithoutFormat,
}: {
  readonly target: ContextMenuTarget
  /** Se o cursor está numa tabela — só então as ações dela aparecem. */
  readonly inTable: boolean
  readonly onTableAction: (action: TableAction) => void
  /** O tipo da lista em que está o cursor, ou `null` fora de lista. */
  readonly inList: 'bulletList' | 'orderedList' | null
  readonly onListAction: (action: ListAction) => void
  readonly onClose: () => void
  readonly onPasteWithoutFormat: () => void
}): React.JSX.Element {
  const showError = useWorkspace((state) => state.showError)
  const readOnly = useWorkspace((state) => state.readOnly)
  const t = useT()

  /** Toda ação fecha o menu — inclusive quando falha, para o erro ficar visível. */
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
            // Um item apagado, e não item nenhum: o menu abre por causa da palavra
            // sublinhada, e sem nada ali pareceria que o menu é que quebrou.
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
          {/* "Ignorar" vale até fechar o aplicativo: o Chromium não tem lista de
              ignorados, então ela é imitada com uma entrada temporária no
              dicionário, desfeita na saída. */}
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

      {/* A numeração vem logo depois da área de transferência, como no Word; e
          reiniciar ou continuar só faz sentido em lista numerada. */}
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

      {/* As ações de tabela vêm depois da área de transferência, como no Word, e
          só com o cursor dentro de uma: fora dela seriam todas itens apagados.
          "Inserir tabela" fica de fora aqui — tabela dentro de tabela é caso de
          quem sabe o que quer, e mora no menu "Tabela". */}
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
