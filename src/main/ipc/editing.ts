import { clipboard } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { EditCommand } from '@shared/types.js'
import { editorPreferences, updatePreferences } from '../preferences.js'
import { rememberWord } from '../spellcheck.js'
import { handle } from './registry.js'

/** Pelo `event.sender`, e não pela janela em foco: com a janela oculta de impressão no ar, o foco é aposta. */
export function registerEditingHandlers(): void {
  handle(IpcChannel.PreferencesGet, () => editorPreferences())
  handle(IpcChannel.PreferencesSet, (payload) => updatePreferences(payload))

  handle(IpcChannel.EditCommandRun, (payload, event) => {
    const contents = event.sender

    // Pelo `webContents`, para o HTML formatado passar pela área de transferência do sistema.
    if (payload.command === EditCommand.Cut) contents.cut()
    if (payload.command === EditCommand.Copy) contents.copy()
    if (payload.command === EditCommand.Paste) contents.paste()

    return { done: true as const }
  })

  // Só o texto: o HTML nem é lido, para não vazar marcação.
  handle(IpcChannel.ClipboardReadText, () => ({
    text: clipboard.readText().slice(0, MAX_TEXT_LENGTH),
  }))

  handle(IpcChannel.SpellReplaceWord, (payload, event) => {
    // Trocar no editor erraria quando a mesma palavra aparece duas vezes na linha.
    event.sender.replaceMisspelling(payload.word)
    return { replaced: true as const }
  })

  handle(IpcChannel.SpellAddWord, (payload, event) => ({
    added: rememberWord(event.sender.session, payload.word, payload.scope),
  }))
}
