import { create } from 'zustand'
import { MenuCommand, RevisionView } from '@shared/types.js'

/**
 * Neither a preference nor the document: as in Word, each window looks its own way, and printing
 * follows.
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

/** `null` if the command is not one of them. */
export function revisionViewOfCommand(command: MenuCommand): RevisionView | null {
  return BY_COMMAND[command] ?? null
}
