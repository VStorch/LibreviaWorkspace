import { MenuCommand, RevisionView } from '@shared/types.js'
import type { CommandHandler, CommandHandlers } from '../commands/execute.js'
import { activeWorkspaceStore } from './tabs.js'

const viewCommand = (view: RevisionView): CommandHandler => ({
  run: () => activeWorkspaceStore().getState().setRevisionView(view),
})

export const REVISION_VIEW_HANDLERS: CommandHandlers = {
  [MenuCommand.ShowAllMarkup]: viewCommand(RevisionView.All),
  [MenuCommand.ShowSimpleMarkup]: viewCommand(RevisionView.Simple),
  [MenuCommand.ShowNoMarkup]: viewCommand(RevisionView.None),
  [MenuCommand.ShowOriginal]: viewCommand(RevisionView.Original),
}
