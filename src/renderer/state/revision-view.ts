import { create } from 'zustand'
import { MenuCommand, RevisionView } from '@shared/types.js'

/** Não é preferência nem documento: como no Word, cada janela olha do seu jeito, e a impressão segue. */
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

/** `null` se o comando não é um deles. */
export function revisionViewOfCommand(command: MenuCommand): RevisionView | null {
  return BY_COMMAND[command] ?? null
}
