import { useRef, useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import { isHiddenBookmark, isValidBookmarkName } from '@services/document/bookmarks.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { bookmarksOf, goToBookmark } from './extensions/bookmark.js'

/**
 * An existing name **moves** the bookmark, as in Word. Hidden ones stay out of the list until the
 * box is checked. When read-only only "Go to" works.
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
  const input = useRef<HTMLInputElement>(null)
  const [showHidden, setShowHidden] = useState(false)
  const [order, setOrder] = useState<BookmarkOrder>('name')

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
    // After the dialog leaves: the unmounted field would take focus with it.
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
          ref={input}
          autoFocus
          maxLength={40}
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') add()
          }}
        />
      </label>

      {name !== '' && !valid && <p className="popover__error">{t('references.bookmark.invalid')}</p>}

      <BookmarkList
        listed={listed}
        selected={name}
        onSelect={setName}
        onOpen={(bookmark) => goToBookmark(editor.view, bookmark)}
      />

      <ListOptions order={order} onOrder={setOrder} showHidden={showHidden} onShowHidden={setShowHidden} />

      {readOnly && <p className="popover__hint">{t('references.bookmark.readOnly')}</p>}

      <BookmarkActions
        canAdd={!readOnly && valid}
        exists={exists}
        readOnly={readOnly}
        onAdd={add}
        onDelete={() => {
          // The dialog stays open, and `Esc` belongs to it.
          editor.commands.deleteBookmark(name)
          setName('')
          // The button greys out, and focus would land on the window body, away from `Esc`.
          input.current?.focus()
        }}
        onGoTo={() => {
          goToBookmark(editor.view, name)
          onClose()
        }}
        onClose={onClose}
      />
    </div>
  )
}

type BookmarkOrder = 'name' | 'location'

function BookmarkList({
  listed,
  selected,
  onSelect,
  onOpen,
}: {
  listed: readonly { readonly name: string }[]
  selected: string
  onSelect: (name: string) => void
  onOpen: (name: string) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <ul className="styles-list" aria-label={t('references.bookmark.list')} role="listbox">
      {listed.length === 0 && <li className="styles-list__empty">{t('references.bookmark.empty')}</li>}
      {listed.map((bookmark) => (
        <li
          key={bookmark.name}
          role="option"
          tabIndex={0}
          aria-selected={bookmark.name === selected}
          className={`styles-list__item${bookmark.name === selected ? ' styles-list__item--selected' : ''}`}
          onClick={() => onSelect(bookmark.name)}
          onDoubleClick={() => onOpen(bookmark.name)}
          onKeyDown={(event) => {
            if (event.key === 'Enter' || event.key === ' ') onSelect(bookmark.name)
          }}
        >
          <span className="styles-list__name">{bookmark.name}</span>
        </li>
      ))}
    </ul>
  )
}

function ListOptions({
  order,
  onOrder,
  showHidden,
  onShowHidden,
}: {
  order: BookmarkOrder
  onOrder: (order: BookmarkOrder) => void
  showHidden: boolean
  onShowHidden: (show: boolean) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('references.bookmark.order')}</span>
          <select value={order} onChange={(event) => onOrder(event.target.value as BookmarkOrder)}>
            <option value="name">{t('references.bookmark.byName')}</option>
            <option value="location">{t('references.bookmark.byLocation')}</option>
          </select>
        </label>
      </div>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={showHidden}
          onChange={(event) => onShowHidden(event.target.checked)}
        />
        {t('references.bookmark.hidden')}
      </label>
    </>
  )
}

function BookmarkActions({
  canAdd,
  exists,
  readOnly,
  onAdd,
  onDelete,
  onGoTo,
  onClose,
}: {
  canAdd: boolean
  exists: boolean
  readOnly: boolean
  onAdd: () => void
  onDelete: () => void
  onGoTo: () => void
  onClose: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__actions">
      <button type="button" className="btn" disabled={!canAdd} onClick={onAdd}>
        {exists ? t('references.bookmark.move') : t('references.bookmark.add')}
      </button>
      <button type="button" className="btn" disabled={readOnly || !exists} onClick={onDelete}>
        {t('references.bookmark.delete')}
      </button>
      <button type="button" className="btn" disabled={!exists} onClick={onGoTo}>
        {t('references.bookmark.goTo')}
      </button>
      <span className="popover__spacer" />
      <button type="button" className="btn btn--primary" onClick={onClose}>
        {t('document.common.close')}
      </button>
    </div>
  )
}
