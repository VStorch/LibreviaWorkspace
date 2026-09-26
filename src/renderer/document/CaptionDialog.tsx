import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../i18n.js'
import { captionLabels, insertCaption, type ReferenceContext } from './references.js'

/**
 * Inserir legenda: o rótulo (Figura, Tabela, Equação ou um que o documento já
 * use), o texto e o lado do bloco do cursor. O número é um campo `SEQ` — ver
 * `insertCaption`.
 */
export function CaptionDialog({
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
    captionLabels(editor.state.doc, [
      t('references.caption.figure'),
      t('references.caption.table'),
      t('references.caption.equation'),
    ]),
  )
  const [label, setLabel] = useState(labels[0] ?? '')
  const [text, setText] = useState('')
  const [above, setAbove] = useState(false)

  function insert(): void {
    insertCaption(editor, context(), { label, text, above })
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('references.caption.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <label className="popover__field">
        <span>{t('references.caption.label')}</span>
        <select value={label} onChange={(event) => setLabel(event.target.value)}>
          {labels.map((item) => (
            <option key={item} value={item}>
              {item}
            </option>
          ))}
        </select>
      </label>
      <label className="popover__field">
        <span>{t('references.caption.text')}</span>
        <input
          type="text"
          value={text}
          autoFocus
          onChange={(event) => setText(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') insert()
          }}
        />
      </label>
      <label className="popover__field">
        <span>{t('references.caption.position')}</span>
        <select
          value={above ? 'above' : 'below'}
          onChange={(event) => setAbove(event.target.value === 'above')}
        >
          <option value="below">{t('references.caption.below')}</option>
          <option value="above">{t('references.caption.above')}</option>
        </select>
      </label>
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" disabled={label.trim() === ''} onClick={insert}>
          {t('references.insert')}
        </button>
      </div>
    </div>
  )
}
