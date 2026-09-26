import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { isHiddenBookmark, isValidBookmarkName } from '@services/document/bookmarks.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { bookmarksOf, goToBookmark } from './extensions/bookmark.js'

/**
 * Marcadores: adicionar, ir para e excluir — o diálogo "Indicador" do Word.
 *
 * O nome digitado que já existe **move** o marcador para a seleção, como lá: o
 * marcador é um só. Os ocultos (`_Toc…`, `_Ref…`) ficam fora da lista até a caixa
 * ser marcada; são do sumário e das referências cruzadas, e apagá-los por engano
 * quebraria as duas coisas.
 *
 * No somente leitura só "Ir para" funciona: andar pelo documento não o muda.
 */
export function BookmarkDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const readOnly = useWorkspace((state) => state.readOnly)
  const [name, setName] = useState('')
  const [showHidden, setShowHidden] = useState(false)
  const [order, setOrder] = useState<'name' | 'location'>('name')

  const bookmarks = useEditorState({
    editor,
    selector: ({ editor: live }) =>
      bookmarksOf(live.state.doc).map(({ name: label, pos }) => ({ name: label, pos })),
  })

  const listed = bookmarks
    .filter((bookmark) => showHidden || !isHiddenBookmark(bookmark.name))
    .sort((left, right) => (order === 'name' ? left.name.localeCompare(right.name) : left.pos - right.pos))
  const exists = bookmarks.some((bookmark) => bookmark.name === name)
  const valid = isValidBookmarkName(name)

  function add(): void {
    if (readOnly || !valid) return
    editor.commands.setBookmark(name)
    onClose()
    // O foco volta ao texto depois de o diálogo sair da tela: devolvido antes,
    // o campo do nome o levava junto ao ser desmontado, e o que se digitava em
    // seguida não ia para lugar nenhum.
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('references.bookmark.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <label className="popover__field">
        <span>{t('references.bookmark.name')}</span>
        <input
          type="text"
          value={name}
          autoFocus
          maxLength={40}
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') add()
          }}
        />
      </label>

      {name !== '' && !valid && <p className="popover__error">{t('references.bookmark.invalid')}</p>}

      <ul className="styles-list" aria-label={t('references.bookmark.list')} role="listbox">
        {listed.length === 0 && <li className="styles-list__empty">{t('references.bookmark.empty')}</li>}
        {listed.map((bookmark) => (
          <li
            key={bookmark.name}
            role="option"
            tabIndex={0}
            aria-selected={bookmark.name === name}
            className={`styles-list__item${bookmark.name === name ? ' styles-list__item--selected' : ''}`}
            onClick={() => setName(bookmark.name)}
            onDoubleClick={() => goToBookmark(editor.view, bookmark.name)}
            onKeyDown={(event) => {
              if (event.key === 'Enter' || event.key === ' ') setName(bookmark.name)
            }}
          >
            <span className="styles-list__name">{bookmark.name}</span>
          </li>
        ))}
      </ul>

      <div className="popover__row">
        <label className="popover__field">
          <span>{t('references.bookmark.order')}</span>
          <select value={order} onChange={(event) => setOrder(event.target.value as 'name' | 'location')}>
            <option value="name">{t('references.bookmark.byName')}</option>
            <option value="location">{t('references.bookmark.byLocation')}</option>
          </select>
        </label>
      </div>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={showHidden}
          onChange={(event) => setShowHidden(event.target.checked)}
        />
        {t('references.bookmark.hidden')}
      </label>

      {readOnly && <p className="popover__hint">{t('references.bookmark.readOnly')}</p>}

      <div className="popover__actions">
        <button type="button" className="btn" disabled={readOnly || !valid} onClick={add}>
          {exists ? t('references.bookmark.move') : t('references.bookmark.add')}
        </button>
        <button
          type="button"
          className="btn"
          disabled={readOnly || !exists}
          onClick={() => {
            editor.chain().focus().deleteBookmark(name).run()
            setName('')
          }}
        >
          {t('references.bookmark.delete')}
        </button>
        <button
          type="button"
          className="btn"
          disabled={!exists}
          onClick={() => {
            goToBookmark(editor.view, name)
            onClose()
          }}
        >
          {t('references.bookmark.goTo')}
        </button>
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" onClick={onClose}>
          {t('document.common.close')}
        </button>
      </div>
    </div>
  )
}
