import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import { columnResizingPluginKey } from '@tiptap/pm/tables'
import { screenScaleOf } from '../screen-scale.js'

/**
 * O `columnResizing` mede o arrasto por `clientX`, que com zoom chega na escala
 * da tela: a 150 %, arrastar 30 px alargaria 45. Durante o gesto, o `clientX` de
 * cada evento vira `clientX / escala`, por ouvintes de captura na janela, que
 * correm antes dos do plugin.
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
