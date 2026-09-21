import { clipboard } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { EditCommand } from '@shared/types.js'
import { editorPreferences, updatePreferences } from '../preferences.js'
import { rememberWord } from '../spellcheck.js'
import { handle } from './registry.js'

/**
 * As operações de edição que dependem de algo que o renderer não alcança: a área
 * de transferência do sistema, o corretor ortográfico e as preferências.
 *
 * Todas usam o `event.sender`, e não a janela em foco: a ação pertence a **quem
 * pediu**. Com a janela oculta de impressão no ar, "a janela em foco" é uma
 * aposta.
 */
export function registerEditingHandlers(): void {
  handle(IpcChannel.PreferencesGet, () => editorPreferences())
  handle(IpcChannel.PreferencesSet, (payload) => updatePreferences(payload))

  handle(IpcChannel.EditCommandRun, (payload, event) => {
    const contents = event.sender

    // Pelo `webContents`, e não por `document.execCommand` no renderer: é o que
    // faz recortar e colar passarem pela área de transferência do sistema, com o
    // HTML formatado, em vez de um atalho que só funciona dentro da página.
    if (payload.command === EditCommand.Cut) contents.cut()
    if (payload.command === EditCommand.Copy) contents.copy()
    if (payload.command === EditCommand.Paste) contents.paste()

    return { done: true as const }
  })

  // Só o texto: é disto que "colar sem formatação" é feito. O HTML da área de
  // transferência nem é lido, para não haver como vazar marcação por engano.
  handle(IpcChannel.ClipboardReadText, () => ({
    text: clipboard.readText().slice(0, MAX_TEXT_LENGTH),
  }))

  handle(IpcChannel.SpellReplaceWord, (payload, event) => {
    // `replaceMisspelling` troca a palavra que o corretor marcou debaixo do
    // cursor. Fazer a troca no editor daria no mesmo em quase todo caso e erraria
    // no que importa: quando a mesma palavra aparece duas vezes na linha.
    event.sender.replaceMisspelling(payload.word)
    return { replaced: true as const }
  })

  handle(IpcChannel.SpellAddWord, (payload, event) => ({
    added: rememberWord(event.sender.session, payload.word, payload.scope),
  }))
}
