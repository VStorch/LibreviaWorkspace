import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import { columnResizingPluginKey } from '@tiptap/pm/tables'
import { screenScaleOf } from '../screen-scale.js'

/**
 * O arrasto da divisória de coluna na escala do documento.
 *
 * O `columnResizing` do prosemirror-tables mede o arrasto por `clientX` —
 * largura nova = largura de partida + (x agora − x de partida) —, e com a folha
 * ampliada por `transform` o deslocamento chega na escala da tela: a 150 %,
 * arrastar 30 px alargava a coluna 30 px de documento, que na tela são 45. O
 * plugin não tem opção para isso, e copiá-lo seria manter um fork.
 *
 * Então os eventos do arrasto chegam a ele já convertidos: durante o gesto, o
 * `clientX` de cada evento é redefinido como `clientX / escala`. Como partida e
 * chegada passam pela mesma conversão, a diferença sai dividida pela escala, que
 * é o que se quer. Fora do arrasto nada muda — a detecção da divisória compara
 * `clientX` com `getBoundingClientRect`, os dois na escala da tela, e continua
 * certa sem ajuda. Os ouvintes são de captura na janela, para correr antes dos
 * do plugin (`mousedown` na visão, `mousemove`/`mouseup` na janela).
 */
export const ZoomedColumnResize = Extension.create({
  name: 'zoomedColumnResize',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        view: (view) => {
          const window = view.dom.ownerDocument.defaultView
          if (window === null) return {}

          const convert = (event: MouseEvent): void => {
            const state = columnResizingPluginKey.getState(view.state)
            if (state === undefined || state === null) return
            const starting =
              event.type === 'mousedown' && state.activeHandle > -1 && view.dom.contains(event.target as Node)
            if (!starting && !state.dragging) return
            const scale = screenScaleOf(view.dom as HTMLElement)
            if (Math.abs(scale - 1) < 0.001) return
            Object.defineProperty(event, 'clientX', { value: event.clientX / scale, configurable: true })
          }

          const options = { capture: true }
          window.addEventListener('mousedown', convert, options)
          window.addEventListener('mousemove', convert, options)
          window.addEventListener('mouseup', convert, options)
          return {
            destroy: () => {
              window.removeEventListener('mousedown', convert, options)
              window.removeEventListener('mousemove', convert, options)
              window.removeEventListener('mouseup', convert, options)
            },
          }
        },
      }),
    ]
  },
})
