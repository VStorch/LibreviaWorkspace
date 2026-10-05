import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { NoteKind } from '@services/document/notes.js'
import { DialogActions } from '../components/DialogActions.js'
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

const NOTE_TYPES: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'note:footnote',
  [NoteKind.Endnote]: 'note:endnote',
}

/** Becomes a `REF`, `NOTEREF` or `PAGEREF` field that F9 recomputes; see `insertCrossReference`. */
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
  const [type, setType] = useState('heading')
  const [key, setKey] = useState<string | null>(null)
  const [show, setShow] = useState<CrossReferenceShow>('text')
  const [link, setLink] = useState(true)

  const kind = kindOf(type)
  const targets = useMemo(
    () => crossReferenceTargets(editor.state.doc, context().styles, kind, noteLabelsOf(editor.state)),
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
      <TargetTypeSelect
        value={type}
        labels={labels}
        onChange={(next) => {
          setType(next)
          setKey(null)
          if (next.startsWith('note:')) setShow(show === 'page' ? 'page' : 'number')
          else if (next !== type && show === 'number') setShow('text')
        }}
      />

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

      <ShowSelect kind={kind} value={show} onChange={setShow} />

      <label className="popover__check">
        <input type="checkbox" checked={link} onChange={(event) => setLink(event.target.checked)} />
        {t('references.crossRef.link')}
      </label>

      <DialogActions
        confirmLabel={t('references.insert')}
        onConfirm={insert}
        onCancel={onClose}
        disabled={chosen === null}
      />
    </div>
  )
}

/** `heading`, `bookmark`, `note:<kind>` or `caption:<label>`: a single value for the picker. */
function kindOf(type: string): CrossReferenceKind {
  if (type === 'heading' || type === 'bookmark') return { type }
  if (type === NOTE_TYPES[NoteKind.Footnote] || type === NOTE_TYPES[NoteKind.Endnote]) {
    return {
      type: 'note',
      kind: type === NOTE_TYPES[NoteKind.Endnote] ? NoteKind.Endnote : NoteKind.Footnote,
    }
  }
  return { type: 'caption', label: type.slice('caption:'.length) }
}

function TargetTypeSelect({
  value,
  labels,
  onChange,
}: {
  value: string
  labels: readonly string[]
  onChange: (type: string) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <label className="popover__field">
      <span>{t('references.crossRef.type')}</span>
      <select value={value} onChange={(event) => onChange(event.target.value)}>
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
  )
}

function ShowSelect({
  kind,
  value,
  onChange,
}: {
  kind: CrossReferenceKind
  value: CrossReferenceShow
  onChange: (show: CrossReferenceShow) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <label className="popover__field">
      <span>{t('references.crossRef.show')}</span>
      <select value={value} onChange={(event) => onChange(event.target.value as CrossReferenceShow)}>
        {kind.type !== 'note' && <option value="text">{t('references.crossRef.showText')}</option>}
        {kind.type === 'caption' && <option value="number">{t('references.crossRef.showNumber')}</option>}
        {kind.type === 'note' && <option value="number">{t('references.crossRef.showNoteNumber')}</option>}
        <option value="page">{t('references.crossRef.showPage')}</option>
      </select>
    </label>
  )
}
