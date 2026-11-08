import { useState } from 'react'
import { useEditorHandlers } from '../commands/useEditorHandlers.js'
import { EditorCommand } from '../document/editor-commands.js'

const AI_PANE_COMMANDS = [EditorCommand.AiAssistant, EditorCommand.AiProviders]

export interface SheetAiPane {
  readonly open: boolean
  readonly configuring: boolean
  readonly close: () => void
  readonly configure: (open: boolean) => void
}

/** The same menu items as in the document: the assistant toggles the panel, the setup opens over it. */
export function useSheetAiPane(): SheetAiPane {
  const [open, setOpen] = useState(false)
  const [configuring, setConfiguring] = useState(false)
  useEditorHandlers(AI_PANE_COMMANDS, (command) => {
    if (command === EditorCommand.AiAssistant) setOpen((current) => !current)
    if (command === EditorCommand.AiProviders) setConfiguring(true)
  })
  return { open, configuring, close: () => setOpen(false), configure: setConfiguring }
}
