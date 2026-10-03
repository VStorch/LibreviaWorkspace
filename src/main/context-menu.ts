import type { WebContents } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { ContextMenuTarget } from '@shared/types.js'
import { sendPush } from './window.js'

/**
 * As sugestões do corretor só existem no evento `context-menu` do `webContents`.
 * O main diz o que havia debaixo do cursor, e o renderer desenha o menu com o
 * componente da planilha: um `Menu.popup()` nativo não ofereceria "colar sem
 * formatação" nem seria clicável pelo Playwright.
 */
export function installContextMenu(contents: WebContents): void {
  contents.on('context-menu', (_event, params) => {
    const target: ContextMenuTarget = {
      // Um clique na borda de uma janela redimensionada chega com fração.
      x: Math.max(0, Math.round(params.x)),
      y: Math.max(0, Math.round(params.y)),
      editable: params.isEditable,
      // Palavra além do limite fica de fora inteira: um pedaço dela iria para o dicionário.
      misspelledWord: params.misspelledWord.length <= 200 ? params.misspelledWord : '',
      // O Chromium manda cinco; o corte protege o menu, e sugestão fora do limite derrubaria o schema.
      dictionarySuggestions: params.dictionarySuggestions
        .filter((suggestion) => suggestion.length <= 200)
        .slice(0, 5),
      canCut: params.editFlags.canCut,
      canCopy: params.editFlags.canCopy,
      canPaste: params.editFlags.canPaste,
    }

    // Ouvinte do Electron, sem `registry` para capturar: uma exceção derrubaria o main.
    try {
      sendPush(contents, IpcChannel.ContextMenuRequested, target)
    } catch (cause) {
      console.error('[context-menu] alvo fora do contrato, menu não enviado:', cause)
    }
  })
}
