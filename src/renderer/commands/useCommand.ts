import { useSyncExternalStore } from 'react'
import { COMMANDS, type CommandId, type CommandMeta } from '@shared/commands/registry.js'
import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { executeCommand, isCommandAvailable, onAvailabilityChanged } from './execute.js'

export interface BoundCommand {
  readonly enabled: boolean
  readonly shortcut: string | undefined
  readonly run: () => void
}

/** The same command as the menu's: it greys out and answers exactly when the menu item does. */
export function useCommand(id: CommandId): BoundCommand {
  const enabled = useSyncExternalStore(onAvailabilityChanged, () => isCommandAvailable(id))
  const meta: CommandMeta = COMMANDS[id]
  return {
    enabled,
    shortcut: meta.shortcut === undefined ? undefined : shortcutHintOf(SHORTCUTS[meta.shortcut]),
    run: () => void executeCommand(id),
  }
}
