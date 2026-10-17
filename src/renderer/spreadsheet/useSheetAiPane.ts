import { useEffect, useState } from 'react'
import { EditorCommand, onEditorCommand } from '../document/editor-commands.js'

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
  useEffect(
    () =>
      onEditorCommand((command) => {
        if (command === EditorCommand.AiAssistant) setOpen((current) => !current)
        if (command === EditorCommand.AiProviders) setConfiguring(true)
      }),
    [],
  )
  return { open, configuring, close: () => setOpen(false), configure: setConfiguring }
}
