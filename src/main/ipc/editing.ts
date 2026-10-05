import { clipboard } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { EditCommand } from '@shared/types.js'
import { editorPreferences, updatePreferences } from '../preferences.js'
import { rememberWord } from '../spellcheck.js'
import { handle } from './registry.js'

/**
 * Through `event.sender`, not the focused window: with the hidden print window around, focus is a
 * guess.
 */
export function registerEditingHandlers(): void {
  handle(IpcChannel.PreferencesGet, () => editorPreferences())
  handle(IpcChannel.PreferencesSet, (payload) => updatePreferences(payload))

  handle(IpcChannel.EditCommandRun, (payload, event) => {
    const contents = event.sender

    // Through `webContents`, so formatted HTML goes through the system clipboard.
    if (payload.command === EditCommand.Cut) contents.cut()
    if (payload.command === EditCommand.Copy) contents.copy()
    if (payload.command === EditCommand.Paste) contents.paste()

    return { done: true as const }
  })

  // Text only: the HTML is not even read, so no markup leaks.
  handle(IpcChannel.ClipboardReadText, () => ({
    text: clipboard.readText().slice(0, MAX_TEXT_LENGTH),
  }))

  handle(IpcChannel.SpellReplaceWord, (payload, event) => {
    // Replacing in the editor would go wrong when the same word appears twice on the line.
    event.sender.replaceMisspelling(payload.word)
    return { replaced: true as const }
  })

  handle(IpcChannel.SpellAddWord, (payload, event) => ({
    added: rememberWord(event.sender.session, payload.word, payload.scope),
  }))
}
