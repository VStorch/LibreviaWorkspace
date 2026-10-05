import type { WebContents } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { ContextMenuTarget } from '@shared/types.js'
import { sendPush } from './window.js'

/**
 * Spellchecker suggestions only exist in the `webContents` `context-menu` event. Main reports what
 * was under the cursor, and the renderer draws the menu with the spreadsheet component: a native
 * `Menu.popup()` would offer neither "paste without formatting" nor be clickable by Playwright.
 */
export function installContextMenu(contents: WebContents): void {
  contents.on('context-menu', (_event, params) => {
    const target: ContextMenuTarget = {
      // A click on the edge of a resized window arrives with a fraction.
      x: Math.max(0, Math.round(params.x)),
      y: Math.max(0, Math.round(params.y)),
      editable: params.isEditable,
      // A word past the limit is dropped whole: a fragment of it would go into the dictionary.
      misspelledWord: params.misspelledWord.length <= 200 ? params.misspelledWord : '',
      // Chromium sends five; the cut protects the menu, and a suggestion past the limit would fail
      // the schema.
      dictionarySuggestions: params.dictionarySuggestions
        .filter((suggestion) => suggestion.length <= 200)
        .slice(0, 5),
      canCut: params.editFlags.canCut,
      canCopy: params.editFlags.canCopy,
      canPaste: params.editFlags.canPaste,
    }

    // An Electron listener, with no `registry` to catch: an exception would bring main down.
    try {
      sendPush(contents, IpcChannel.ContextMenuRequested, target)
    } catch (cause) {
      console.error('[context-menu] alvo fora do contrato, menu não enviado:', cause)
    }
  })
}
