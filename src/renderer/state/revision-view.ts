import { create } from 'zustand'
import { MenuCommand, RevisionView } from '@shared/types.js'

/**
 * Como esta janela mostra as alterações controladas (Revisão → Mostrar).
 *
 * Não é preferência nem documento: no Word cada janela olha o arquivo do seu
 * jeito, e gravar a escolha faria a próxima abertura esconder alterações sem a
 * pessoa saber que elas estão lá. Mora aqui, como o zoom de "ajustar à largura";
 * o editor a aplica (ver extensions/revision-view.ts) e a impressão a segue.
 */
export const useRevisionView = create<{ view: RevisionView }>(() => ({ view: RevisionView.All }))

export function setRevisionView(view: RevisionView): void {
  if (useRevisionView.getState().view !== view) useRevisionView.setState({ view })
}

const BY_COMMAND: Partial<Record<MenuCommand, RevisionView>> = {
  [MenuCommand.ShowAllMarkup]: RevisionView.All,
  [MenuCommand.ShowSimpleMarkup]: RevisionView.Simple,
  [MenuCommand.ShowNoMarkup]: RevisionView.None,
  [MenuCommand.ShowOriginal]: RevisionView.Original,
}

/** O comando do menu que escolhe o jeito de mostrar — ou `null`, se não é um deles. */
export function revisionViewOfCommand(command: MenuCommand): RevisionView | null {
  return BY_COMMAND[command] ?? null
}
