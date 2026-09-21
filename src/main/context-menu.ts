import type { WebContents } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { ContextMenuTarget } from '@shared/types.js'
import { sendPush } from './window.js'

/**
 * O clique com o botão direito, do main para o renderer.
 *
 * ## Por que o main entra nisto
 *
 * Um menu de contexto de editor de texto se desenharia todo no renderer — o da
 * planilha se desenha — se não fosse um detalhe: as sugestões do corretor
 * ortográfico **só existem aqui**. O Chromium as calcula ao montar o evento
 * `context-menu` do `webContents`, e não há API no renderer que as alcance.
 *
 * Então a divisão é esta: o main diz o que havia debaixo do cursor, o renderer
 * desenha o menu (com o mesmo componente da planilha) e volta pelo IPC para pedir
 * as ações que dependem do `webContents` — recortar, colar, trocar a palavra.
 *
 * Um `Menu.popup()` nativo daqui seria menos código e pior: teria a aparência do
 * sistema em vez da do aplicativo, não conseguiria oferecer "colar sem
 * formatação" (que precisa do editor) e, em teste, nenhum Playwright o clica.
 *
 * O evento não é cancelado porque não há o que cancelar: o Electron não desenha
 * menu de contexto por conta própria.
 */
export function installContextMenu(contents: WebContents): void {
  contents.on('context-menu', (_event, params) => {
    const target: ContextMenuTarget = {
      // Inteiros e nunca negativos: o schema recusa o resto, e um clique na
      // borda de uma janela redimensionada chega com fração.
      x: Math.max(0, Math.round(params.x)),
      y: Math.max(0, Math.round(params.y)),
      editable: params.isEditable,
      // Cortada no limite do schema, e não truncada: uma "palavra" de duzentos
      // caracteres não tem sugestão nenhuma, e um pedaço dela é o que iria para o
      // dicionário do usuário se ele clicasse em "Adicionar".
      misspelledWord: params.misspelledWord.length <= 200 ? params.misspelledWord : '',
      // Cinco é o que o Chromium manda; o corte protege o menu de crescer além
      // da tela se isso mudar. O filtro de comprimento é do mesmo tipo do de cima:
      // uma sugestão fora do limite do schema levaria o menu inteiro embora.
      dictionarySuggestions: params.dictionarySuggestions
        .filter((suggestion) => suggestion.length <= 200)
        .slice(0, 5),
      canCut: params.editFlags.canCut,
      canCopy: params.editFlags.canCopy,
      canPaste: params.editFlags.canPaste,
    }

    // `sendPush` valida com `.parse` e joga, e aqui não há `registry` para
    // capturar: este é um ouvinte de evento do Electron, e uma exceção daqui é
    // exceção não tratada no processo main. Ficar sem menu de contexto por um
    // clique é ruim; derrubar o main com o documento aberto é pior.
    try {
      sendPush(contents, IpcChannel.ContextMenuRequested, target)
    } catch (cause) {
      console.error('[context-menu] alvo fora do contrato, menu não enviado:', cause)
    }
  })
}
