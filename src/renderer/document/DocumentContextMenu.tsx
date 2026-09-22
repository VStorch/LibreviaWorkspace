import { Fragment } from 'react'
import type { IpcResult } from '@shared/ipc.js'
import { TABLE_ACTIONS, type TableAction } from '@shared/table-actions.js'
import { DictionaryScope, EditCommand, type ContextMenuTarget } from '@shared/types.js'
import { ContextMenu, ContextMenuItem, ContextMenuSeparator } from '../components/ContextMenu.js'
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
export function DocumentContextMenu({
  target,
  inTable,
  onTableAction,
  onClose,
  onPasteWithoutFormat,
}: {
  readonly target: ContextMenuTarget
  /** Se o cursor está numa tabela — só então as ações dela aparecem. */
  readonly inTable: boolean
  readonly onTableAction: (action: TableAction) => void
  readonly onClose: () => void
  readonly onPasteWithoutFormat: () => void
}): React.JSX.Element {
  const showError = useWorkspace((state) => state.showError)
  const readOnly = useWorkspace((state) => state.readOnly)

  /** Toda ação fecha o menu — inclusive quando falha, para o erro ficar visível. */
  const act = (run: () => Promise<IpcResult<unknown>>) => () => {
    onClose()
    void run().then((result) => {
      if (!result.ok) showError(result.error)
    })
  }

  const misspelled = target.misspelledWord !== ''

  return (
    <ContextMenu position={target} label="Ações do documento" onClose={onClose}>
      {misspelled && (
        <>
          {target.dictionarySuggestions.length === 0 ? (
            // Um item apagado, e não item nenhum: o menu abre por causa da palavra
            // sublinhada, e sem nada ali pareceria que o menu é que quebrou.
            <ContextMenuItem disabled onClick={onClose}>
              Nenhuma sugestão
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
            Adicionar ao dicionário
          </ContextMenuItem>
          {/* "Ignorar" vale até fechar o aplicativo: o Chromium não tem lista de
              ignorados, então ela é imitada com uma entrada temporária no
              dicionário, desfeita na saída. */}
          <ContextMenuItem
            onClick={act(() =>
              window.api.spell.addWord({ word: target.misspelledWord, scope: DictionaryScope.Session }),
            )}
          >
            Ignorar nesta sessão
          </ContextMenuItem>

          <ContextMenuSeparator />
        </>
      )}

      <ContextMenuItem
        disabled={!target.canCut || readOnly}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Cut }))}
      >
        Recortar
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canCopy}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Copy }))}
      >
        Copiar
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canPaste || readOnly}
        onClick={act(() => window.api.edit.run({ command: EditCommand.Paste }))}
      >
        Colar
      </ContextMenuItem>
      <ContextMenuItem
        disabled={!target.canPaste || readOnly}
        onClick={() => {
          onClose()
          onPasteWithoutFormat()
        }}
      >
        Colar sem formatação
      </ContextMenuItem>

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
                {action.label}
              </ContextMenuItem>
            </Fragment>
          ))}
        </>
      )}
    </ContextMenu>
  )
}
