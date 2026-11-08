import { useEffect, useRef } from 'react'
import type { EditorCommand } from '../document/editor-commands.js'
import { useOwnTabId } from '../state/tabs.js'
import { editorSource, registerHandlers, type CommandHandlers } from './execute.js'

/**
 * While mounted, `commands` reach `run` only when this view's tab is the active one. `run` may change
 * on every render; the registration stays.
 */
export function useEditorHandlers(
  commands: readonly EditorCommand[],
  run: (command: EditorCommand) => void,
): void {
  const tabId = useOwnTabId()
  const latest = useRef(run)
  latest.current = run

  useEffect(() => {
    const map: CommandHandlers = Object.fromEntries(
      commands.map((command) => [command, { run: () => latest.current(command) }]),
    )
    return registerHandlers(editorSource(tabId), map)
  }, [tabId, commands])
}
