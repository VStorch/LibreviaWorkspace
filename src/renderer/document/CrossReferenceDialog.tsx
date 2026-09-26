import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../i18n.js'
import {
  captionLabels,
  crossReferenceTargets,
  insertCrossReference,
  type CrossReferenceKind,
  type CrossReferenceShow,
  type ReferenceContext,
} from './references.js'

/**
 * Referência cruzada: a um título, a um marcador ou a uma legenda, mostrando o
 * texto, o número (da legenda) ou a página. Vira um campo `REF` ou `PAGEREF` que
 * "Atualizar campos" (F9) recalcula — ver `insertCrossReference`.
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
  // `heading`, `bookmark` ou `caption:<rótulo>` — um valor só para o seletor.
  const [type, setType] = useState('heading')
  const [key, setKey] = useState<string | null>(null)
  const [show, setShow] = useState<CrossReferenceShow>('text')
  const [link, setLink] = useState(true)

  const kind: CrossReferenceKind =
    type === 'heading' || type === 'bookmark'
      ? { type }
      : { type: 'caption', label: type.slice('caption:'.length) }
  const targets = useMemo(
    () => crossReferenceTargets(editor.state.doc, context().styles, kind),
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
            setType(event.target.value)
            setKey(null)
            if (event.target.value !== type && show === 'number') setShow('text')
          }}
        >
          <option value="heading">{t('references.crossRef.heading')}</option>
          <option value="bookmark">{t('references.crossRef.bookmark')}</option>
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
          <option value="text">{t('references.crossRef.showText')}</option>
          {kind.type === 'caption' && <option value="number">{t('references.crossRef.showNumber')}</option>}
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
