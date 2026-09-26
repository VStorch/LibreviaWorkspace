import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { isHiddenBookmark } from '@services/document/bookmarks.js'
import { normalizeLinkUrl } from '@services/document/link.js'
import { outlineOf } from '@services/document/outline.js'
import { useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'
import { bookmarksOf, ensureBlockBookmark } from '../extensions/bookmark.js'
import { outlineBlocksOf } from '../outline-blocks.js'

interface LinkDialogProps {
  readonly editor: Editor
  readonly onClose: () => void
}

/**
 * O destino escolhido em "Lugar neste documento": um marcador pelo nome, ou um
 * título pela posição — o título ganha um marcador oculto na hora de aplicar.
 */
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

  // Os lugares são lidos uma vez, ao abrir: o diálogo não muda o documento
  // enquanto está aberto, e lê-los ao vivo só custaria.
  const [places] = useState(() => {
    const doc = editor.state.doc
    return {
      bookmarks: bookmarksOf(doc)
        .map((bookmark) => bookmark.name)
        // O oculto só quando é o destino do link em edição: sem ele na lista, o
        // seletor abriria em branco num link que funciona.
        .filter((name) => !isHiddenBookmark(name) || (internal && name === existing.slice(1)))
        .sort((left, right) => left.localeCompare(right)),
      headings: outlineOf(outlineBlocksOf(doc), sheet),
    }
  })
  const [place, setPlace] = useState<string>(internal ? `b:${existing.slice(1)}` : '')

  function applyPlace(): void {
    const chosen: Place | null = place.startsWith('b:')
      ? { kind: 'bookmark', name: place.slice(2) }
      : place.startsWith('h:')
        ? { kind: 'heading', pos: Number(place.slice(2)) }
        : null
    if (chosen === null) return

    // O título não tem nome a citar: ganha um marcador oculto em volta do texto,
    // que é o que o Word faz. A posição dos lugares foi lida ao abrir, e o
    // marcador novo entra depois dela — a seleção é mapeada pela transação.
    const label =
      chosen.kind === 'bookmark'
        ? chosen.name
        : (places.headings.find((heading) => heading.pos === chosen.pos)?.text ?? '')
    const name =
      chosen.kind === 'bookmark' ? chosen.name : ensureBlockBookmark(editor.view, chosen.pos, '_Ref')
    if (name === null) return

    const href = `#${name}`
    const chain = editor.chain().focus()
    if (editor.state.selection.empty && !editor.isActive('link')) {
      // Sem texto selecionado, o link leva o nome do lugar, como no Word.
      chain.insertContent({ type: 'text', text: label, marks: [{ type: 'link', attrs: { href } }] }).run()
    } else {
      chain.extendMarkRange('link').setLink({ href }).run()
    }
    onClose()
  }

  function apply(): void {
    if (place !== '') {
      applyPlace()
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

      {/* O lugar neste documento substitui o endereço: um link é uma coisa ou a
          outra, e o campo se apaga para dizer isso. */}
      <label className="popover__field">
        <span>{t('references.link.place')}</span>
        <select value={place} onChange={(event) => setPlace(event.target.value)}>
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

      {rejected && <p className="popover__error">{t('document.linkDialog.invalidAddress')}</p>}

      <div className="popover__actions">
        {existing !== '' && (
          <button type="button" className="btn" onClick={remove}>
            {t('document.linkDialog.remove')}
          </button>
        )}
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={apply}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
