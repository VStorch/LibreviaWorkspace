import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { propertyFieldChoices, propertyFieldResult } from '@services/document/property-fields.js'
import { DialogActions } from '../components/DialogActions.js'
import { useT } from '../i18n.js'
import { insertPropertyField, type ReferenceContext } from './references.js'

/** Becomes a `TITLE`, `AUTHOR` or `DOCPROPERTY` field that F9 and File → Properties update. */
export function PropertyFieldDialog({
  editor,
  context,
  onClose,
}: {
  readonly editor: Editor
  readonly context: () => ReferenceContext
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [{ properties, locale }] = useState(() => context())
  const [choices] = useState(() => propertyFieldChoices(properties))
  const [chosen, setChosen] = useState(0)
  const choice = choices[chosen]
  const value = choice === undefined ? '' : (propertyFieldResult(choice.instr, properties, locale) ?? '')

  function insert(): void {
    if (choice === undefined) return
    insertPropertyField(editor, context(), choice.instr)
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('references.property.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <label className="popover__field">
        <span>{t('references.property.name')}</span>
        <select size={8} value={chosen} onChange={(event) => setChosen(Number(event.target.value))}>
          {choices.map((option, index) => (
            <option key={option.instr} value={index}>
              {option.builtIn ? t(`document.properties.${option.key}`) : option.name}
            </option>
          ))}
        </select>
      </label>
      <p className="popover__hint">
        {t('references.property.value')}: {value === '' ? t('references.property.emptyValue') : value}
      </p>
      <p className="popover__hint">{t('references.property.hint')}</p>
      <DialogActions
        confirmLabel={t('references.insert')}
        onConfirm={insert}
        onCancel={onClose}
        disabled={choice === undefined}
      />
    </div>
  )
}
