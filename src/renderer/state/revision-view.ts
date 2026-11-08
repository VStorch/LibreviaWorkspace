import { create } from 'zustand'
import { MenuCommand, RevisionView } from '@shared/types.js'
import type { CommandHandler, CommandHandlers } from '../commands/execute.js'

/**
 * Neither a preference nor the document: as in Word, each window looks its own way, and printing
 * follows.
 */
export const useRevisionView = create<{ view: RevisionView }>(() => ({ view: RevisionView.All }))

export function setRevisionView(view: RevisionView): void {
  if (useRevisionView.getState().view !== view) useRevisionView.setState({ view })
}

const viewCommand = (view: RevisionView): CommandHandler => ({ run: () => setRevisionView(view) })

export const REVISION_VIEW_HANDLERS: CommandHandlers = {
  [MenuCommand.ShowAllMarkup]: viewCommand(RevisionView.All),
  [MenuCommand.ShowSimpleMarkup]: viewCommand(RevisionView.Simple),
  [MenuCommand.ShowNoMarkup]: viewCommand(RevisionView.None),
  [MenuCommand.ShowOriginal]: viewCommand(RevisionView.Original),
}
