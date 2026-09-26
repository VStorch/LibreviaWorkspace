import { useEditorState, type Editor } from '@tiptap/react'
import { currentEntryIndex, outlineOf } from '@services/document/outline.js'
import { useT } from '../i18n.js'
import { setPreference } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { outlineBlocksOf } from './outline-blocks.js'

/**
 * O painel de navegação: os títulos do documento, e o da seção do cursor em
 * destaque.
 *
 * Só lê. Clicar num título move o cursor para ele e rola a folha até lá — mover
 * o cursor não muda o documento, e por isso o painel funciona igual no somente
 * leitura, que é justamente onde se abre um documento longo para consultar.
 *
 * Ao vivo, como o painel de estilos: a lista acompanha cada título digitado, e o
 * destaque, cada movimento do cursor. O nível é o efetivo — ver `outline.ts`.
 */
export function NavigationPane({ editor }: { readonly editor: Editor }): React.JSX.Element {
  const sheet = useWorkspace((state) => state.styles)
  const t = useT()

  const { entries, current } = useEditorState({
    editor,
    selector: ({ editor: live }) => {
      const found = outlineOf(outlineBlocksOf(live.state.doc), sheet)
      return { entries: found, current: currentEntryIndex(found, live.state.selection.from) }
    },
  })

  function go(pos: number): void {
    // O cursor no começo do texto do título, e a folha rolada até ele. A rolagem
    // é do elemento, e não do `scrollIntoView` do ProseMirror: aquele só garante
    // que o cursor apareça, e o título ficava colado no pé da janela.
    //
    // O foco vai direto à visão, e não pelo comando `focus` do Tiptap: aquele
    // espera o próximo quadro, e a tecla digitada logo depois do clique caía no
    // botão do painel.
    editor.commands.setTextSelection(pos + 1)
    editor.view.focus()
    const dom = editor.view.nodeDOM(pos)
    if (dom instanceof HTMLElement) dom.scrollIntoView({ block: 'start' })
  }

  return (
    <nav className="nav-pane" aria-label={t('references.nav.title')}>
      <div className="nav-pane__header">
        <span className="nav-pane__title">{t('references.nav.title')}</span>
        <button
          type="button"
          className="nav-pane__close"
          aria-label={t('document.common.close')}
          title={t('document.common.close')}
          onClick={() => void setPreference({ navigationPane: false })}
        >
          ×
        </button>
      </div>

      {entries.length === 0 ? (
        <p className="nav-pane__empty">{t('references.nav.empty')}</p>
      ) : (
        <ul className="nav-pane__list">
          {entries.map((entry, index) => (
            <li key={`${entry.pos}`}>
              <button
                type="button"
                className={`nav-pane__entry${index === current ? ' nav-pane__entry--current' : ''}`}
                // O recuo diz o nível; o leitor de tela ouve o mesmo pelo rótulo.
                style={{ paddingLeft: `${8 + (entry.level - 1) * 14}px` }}
                aria-current={index === current ? 'location' : undefined}
                aria-label={t('references.nav.entry', { level: entry.level, text: entry.text })}
                title={entry.text}
                onClick={() => go(entry.pos)}
              >
                {entry.text}
              </button>
            </li>
          ))}
        </ul>
      )}
    </nav>
  )
}
