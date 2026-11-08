import { MenuCommand, RevisionView } from '@shared/types.js'
import type { CommandHandlers } from '../commands/execute.js'
import { activeWorkspaceStore } from './tabs.js'

export const REVISION_VIEW_COMMANDS = {
  [RevisionView.All]: MenuCommand.ShowAllMarkup,
  [RevisionView.Simple]: MenuCommand.ShowSimpleMarkup,
  [RevisionView.None]: MenuCommand.ShowNoMarkup,
  [RevisionView.Original]: MenuCommand.ShowOriginal,
} as const satisfies Record<RevisionView, MenuCommand>

export const REVISION_VIEW_HANDLERS: CommandHandlers = Object.fromEntries(
  Object.entries(REVISION_VIEW_COMMANDS).map(([view, command]) => [
    command,
    {
      run: () =>
        activeWorkspaceStore()
          .getState()
          .setRevisionView(view as RevisionView),
    },
  ]),
)
