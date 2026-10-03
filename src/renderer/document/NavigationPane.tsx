import { useEditorState, type Editor } from '@tiptap/react'
import { currentEntryIndex, outlineOf } from '@services/document/outline.js'
import { useT } from '../i18n.js'
import { setPreference } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { textStartOf } from './extensions/zero-width.js'
import { outlineBlocksOf } from './outline-blocks.js'

/**
 * Clicar move o cursor e rola a folha: não muda o documento, e por isso vale no
 * somente leitura. O nível é o efetivo (`outline.ts`).
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
    // A rolagem é do elemento: o `scrollIntoView` do ProseMirror deixaria o título
    // no pé da janela. O foco vai direto à visão, porque o `focus` do Tiptap espera
    // um quadro. O cursor vai depois das âncoras do começo do título.
    const start = textStartOf(editor.state.doc, pos)
    // O foco antes da seleção, para o cursor do navegador ir junto.
    editor.view.focus()
    editor.commands.setTextSelection(start)
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
