import { clipboard, type WebContents } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { EditCommand } from '@shared/types.js'
import { editorPreferences, updatePreferences } from '../preferences.js'
import { rememberWord } from '../spellcheck.js'
import { handle } from './registry.js'

/**
 * Through `webContents`, so formatted HTML goes through the system clipboard and undo reaches the
 * focused field or editor as Chromium's own (ProseMirror hears it as `historyUndo`).
 */
const RUN: Record<EditCommand, (contents: WebContents) => void> = {
  [EditCommand.Cut]: (contents) => contents.cut(),
  [EditCommand.Copy]: (contents) => contents.copy(),
  [EditCommand.Paste]: (contents) => contents.paste(),
  [EditCommand.Undo]: (contents) => contents.undo(),
  [EditCommand.Redo]: (contents) => contents.redo(),
}

/**
 * Through `event.sender`, not the focused window: with the hidden print window around, focus is a
 * guess.
 */
export function registerEditingHandlers(): void {
  handle(IpcChannel.PreferencesGet, () => editorPreferences())
  handle(IpcChannel.PreferencesSet, (payload) => updatePreferences(payload))

  handle(IpcChannel.EditCommandRun, (payload, event) => {
    RUN[payload.command](event.sender)
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
