import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { NoteKind } from '@services/document/notes.js'
import { useT } from '../i18n.js'
import { noteLabelsOf } from './extensions/note-ref.js'
import {
  captionLabels,
  crossReferenceTargets,
  insertCrossReference,
  type CrossReferenceKind,
  type CrossReferenceShow,
  type ReferenceContext,
} from './references.js'

/** O valor do seletor de tipo para cada tipo de nota. */
const NOTE_TYPES: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'note:footnote',
  [NoteKind.Endnote]: 'note:endnote',
}

/**
 * Referência cruzada: a um título, a um marcador, a uma legenda ou a uma nota,
 * mostrando o texto, o número (da legenda ou da nota) ou a página. Vira um campo
 * `REF`, `NOTEREF` ou `PAGEREF` que "Atualizar campos" (F9) recalcula — ver
 * `insertCrossReference`.
 */
export function CrossReferenceDialog({
  editor,
  context,
  onClose,
}: {
  readonly editor: Editor
  readonly context: () => ReferenceContext
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [labels] = useState(() =>
    captionLabels(editor.state.doc, [t('references.caption.figure'), t('references.caption.table')]),
  )
  // `heading`, `bookmark`, `note:<tipo>` ou `caption:<rótulo>` — um valor só para o seletor.
  const [type, setType] = useState('heading')
  const [key, setKey] = useState<string | null>(null)
  const [show, setShow] = useState<CrossReferenceShow>('text')
  const [link, setLink] = useState(true)

  const kind: CrossReferenceKind =
    type === 'heading' || type === 'bookmark'
      ? { type }
      : type === NOTE_TYPES[NoteKind.Footnote] || type === NOTE_TYPES[NoteKind.Endnote]
        ? { type: 'note', kind: type === NOTE_TYPES[NoteKind.Endnote] ? NoteKind.Endnote : NoteKind.Footnote }
        : { type: 'caption', label: type.slice('caption:'.length) }
  const targets = useMemo(
    () => crossReferenceTargets(editor.state.doc, context().styles, kind, noteLabelsOf(editor.state)),
    // O documento não muda com o diálogo aberto; o tipo sim.
    [type],
  )
  const chosen = key ?? targets[0]?.key ?? null

  function insert(): void {
    if (chosen === null) return
    insertCrossReference(editor, context(), { kind, key: chosen, show, link })
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('references.crossRef.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <label className="popover__field">
        <span>{t('references.crossRef.type')}</span>
        <select
          value={type}
          onChange={(event) => {
            const next = event.target.value
            setType(next)
            setKey(null)
            // A nota mostra o número dela ou a página; o resto, o texto de saída.
            if (next.startsWith('note:')) setShow(show === 'page' ? 'page' : 'number')
            else if (next !== type && show === 'number') setShow('text')
          }}
        >
          <option value="heading">{t('references.crossRef.heading')}</option>
          <option value="bookmark">{t('references.crossRef.bookmark')}</option>
          <option value={NOTE_TYPES[NoteKind.Footnote]}>{t('references.crossRef.footnote')}</option>
          <option value={NOTE_TYPES[NoteKind.Endnote]}>{t('references.crossRef.endnote')}</option>
          {labels.map((label) => (
            <option key={label} value={`caption:${label}`}>
              {label}
            </option>
          ))}
        </select>
      </label>

      <label className="popover__field">
        <span>{t('references.crossRef.target')}</span>
        <select size={8} value={chosen ?? ''} onChange={(event) => setKey(event.target.value)}>
          {targets.map((target) => (
            <option key={target.key} value={target.key}>
              {target.text}
            </option>
          ))}
        </select>
      </label>
      {targets.length === 0 && <p className="popover__hint">{t('references.crossRef.empty')}</p>}

      <label className="popover__field">
        <span>{t('references.crossRef.show')}</span>
        <select value={show} onChange={(event) => setShow(event.target.value as CrossReferenceShow)}>
          {kind.type !== 'note' && <option value="text">{t('references.crossRef.showText')}</option>}
          {kind.type === 'caption' && <option value="number">{t('references.crossRef.showNumber')}</option>}
          {kind.type === 'note' && <option value="number">{t('references.crossRef.showNoteNumber')}</option>}
          <option value="page">{t('references.crossRef.showPage')}</option>
        </select>
      </label>

      <label className="popover__check">
        <input type="checkbox" checked={link} onChange={(event) => setLink(event.target.checked)} />
        {t('references.crossRef.link')}
      </label>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" disabled={chosen === null} onClick={insert}>
          {t('references.insert')}
        </button>
      </div>
    </div>
  )
}
