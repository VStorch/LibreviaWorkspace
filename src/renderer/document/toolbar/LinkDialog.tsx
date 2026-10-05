import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { isHiddenBookmark } from '@services/document/bookmarks.js'
import { normalizeLinkUrl } from '@services/document/link.js'
import { outlineOf } from '@services/document/outline.js'
import { DialogActions } from '../../components/DialogActions.js'
import { useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'
import type { StyleSheet } from '@services/document/styles.js'
import { bookmarksOf, ensureBlockBookmark } from '../extensions/bookmark.js'
import { outlineBlocksOf } from '../outline-blocks.js'

interface LinkDialogProps {
  readonly editor: Editor
  readonly onClose: () => void
}

type Place =
  { readonly kind: 'bookmark'; readonly name: string } | { readonly kind: 'heading'; readonly pos: number }

function placeKey(place: Place): string {
  return place.kind === 'bookmark' ? `b:${place.name}` : `h:${place.pos}`
}

export function LinkDialog({ editor, onClose }: LinkDialogProps): React.JSX.Element {
  const t = useT()
  const sheet = useWorkspace((state) => state.styles)
  const existing = String(editor.getAttributes('link')['href'] ?? '')
  const internal = existing.startsWith('#')
  const [value, setValue] = useState(internal ? '' : existing)
  const [rejected, setRejected] = useState(false)

  const [places] = useState(() => placesOf(editor, sheet, internal ? existing.slice(1) : null))
  const [place, setPlace] = useState<string>(internal ? `b:${existing.slice(1)}` : '')

  function apply(): void {
    if (place !== '') {
      if (linkToPlace(editor, place, places.headings)) onClose()
      return
    }

    const normalized = normalizeLinkUrl(value)
    if (normalized === null) {
      setRejected(true)
      return
    }

    editor.chain().focus().extendMarkRange('link').setLink({ href: normalized }).run()
    onClose()
  }

  function remove(): void {
    editor.chain().focus().extendMarkRange('link').unsetLink().run()
    onClose()
  }

  return (
    <div className="popover" role="dialog" aria-label={t('document.linkDialog.title')}>
      <label className="popover__field">
        <span>{t('document.linkDialog.address')}</span>
        <input
          type="text"
          value={value}
          autoFocus
          disabled={place !== ''}
          placeholder={t('document.linkDialog.placeholder')}
          onChange={(event) => {
            setValue(event.target.value)
            setRejected(false)
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') apply()
            if (event.key === 'Escape') onClose()
          }}
        />
      </label>
      {/* A link is an address or a place: the field greys out to say so. */}
      <PlaceSelect places={places} value={place} onChange={setPlace} />
      {rejected && <p className="popover__error">{t('document.linkDialog.invalidAddress')}</p>}
      <DialogActions confirmLabel={t('document.common.apply')} onConfirm={apply} onCancel={onClose}>
        {existing !== '' && (
          <button type="button" className="btn" onClick={remove}>
            {t('document.linkDialog.remove')}
          </button>
        )}
      </DialogActions>
    </div>
  )
}

interface Places {
  readonly bookmarks: readonly string[]
  readonly headings: ReturnType<typeof outlineOf>
}

/**
 * Read once, on open: the document does not change while the dialog is open. The hidden bookmark only appears
 * when it is the target of the link being edited, otherwise the picker would open blank.
 */
function placesOf(editor: Editor, sheet: StyleSheet, target: string | null): Places {
  const doc = editor.state.doc
  return {
    bookmarks: bookmarksOf(doc)
      .map((bookmark) => bookmark.name)
      .filter((name) => !isHiddenBookmark(name) || name === target)
      .sort((left, right) => left.localeCompare(right)),
    headings: outlineOf(outlineBlocksOf(doc), sheet),
  }
}

/** `false` when the place no longer exists. */
function linkToPlace(editor: Editor, place: string, headings: Places['headings']): boolean {
  const chosen: Place | null = place.startsWith('b:')
    ? { kind: 'bookmark', name: place.slice(2) }
    : place.startsWith('h:')
      ? { kind: 'heading', pos: Number(place.slice(2)) }
      : null
  if (chosen === null) return false

  // The heading gets a hidden bookmark, as in Word; the transaction maps the selection.
  const label =
    chosen.kind === 'bookmark'
      ? chosen.name
      : (headings.find((heading) => heading.pos === chosen.pos)?.text ?? '')
  const name = chosen.kind === 'bookmark' ? chosen.name : ensureBlockBookmark(editor.view, chosen.pos, '_Ref')
  if (name === null) return false

  const href = `#${name}`
  const chain = editor.chain().focus()
  if (editor.state.selection.empty && !editor.isActive('link')) {
    // With no text selected, the link takes the name of the place, as in Word.
    chain.insertContent({ type: 'text', text: label, marks: [{ type: 'link', attrs: { href } }] }).run()
  } else {
    chain.extendMarkRange('link').setLink({ href }).run()
  }
  return true
}

function PlaceSelect({
  places,
  value,
  onChange,
}: {
  places: Places
  value: string
  onChange: (place: string) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <label className="popover__field">
      <span>{t('references.link.place')}</span>
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">{t('references.link.noPlace')}</option>
        {places.headings.length > 0 && (
          <optgroup label={t('references.link.headings')}>
            {places.headings.map((heading) => (
              <option key={placeKey({ kind: 'heading', pos: heading.pos })} value={`h:${heading.pos}`}>
                {`${' '.repeat(heading.level - 1)}${heading.text}`}
              </option>
            ))}
          </optgroup>
        )}
        {places.bookmarks.length > 0 && (
          <optgroup label={t('references.link.bookmarks')}>
            {places.bookmarks.map((name) => (
              <option key={name} value={`b:${name}`}>
                {name}
              </option>
            ))}
          </optgroup>
        )}
      </select>
    </label>
  )
}
