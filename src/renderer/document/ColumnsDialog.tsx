import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { allSections, withColumns } from '@services/document/sections.js'
import { useT } from '../i18n.js'
import { commitSections, resolvedOf, sectionAtCursor } from './section-commands.js'

/**
 * Formatar → Colunas: quantas, o espaço entre elas e a linha separadora, na
 * seção do cursor ou no documento todo — como a configuração de página.
 */
export function ColumnsDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  // As seções que o texto usa, na ordem dele — ver `resolveSections`.
  const [{ page, sections, bodyId }] = useState(() => resolvedOf(editor.state.doc))
  const [index] = useState(() => sectionAtCursor(editor))
  const current = allSections(page, sections)[index] ?? page
  const [count, setCount] = useState(current.columns?.count ?? 1)
  const [spaceMm, setSpaceMm] = useState(current.columns?.spaceMm ?? 12.7)
  const [separator, setSeparator] = useState(current.columns?.separator ?? false)
  const [scope, setScope] = useState<'section' | 'document'>('section')

  const valid = Number.isInteger(count) && count >= 1 && count <= 12 && spaceMm >= 0 && spaceMm <= 100

  function apply(): void {
    if (!valid) return
    commitSections(withColumns({ page, sections }, index, { count, spaceMm, separator }, scope), bodyId)
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('document.columns.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.columns.count')}</span>
          <input
            type="number"
            min={1}
            max={12}
            step={1}
            value={count}
            autoFocus
            onChange={(event) => setCount(Math.round(Number(event.target.value)))}
          />
        </label>
        <label className="popover__field popover__field--narrow">
          <span>{t('document.columns.spacing')}</span>
          <input
            type="number"
            min={0}
            max={100}
            step={0.5}
            value={spaceMm}
            disabled={count <= 1}
            onChange={(event) => setSpaceMm(Number(event.target.value))}
          />
        </label>
      </div>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={separator}
          disabled={count <= 1}
          onChange={(event) => setSeparator(event.target.checked)}
        />
        {t('document.columns.separator')}
      </label>
      {sections.length > 0 && (
        <div className="popover__row">
          {(['section', 'document'] as const).map((choice) => (
            <label key={choice} className="popover__check">
              <input
                type="radio"
                name="columns-scope"
                checked={scope === choice}
                onChange={() => setScope(choice)}
              />
              {t(
                choice === 'section' ? 'document.pageSetup.thisSection' : 'document.pageSetup.wholeDocument',
              )}
            </label>
          ))}
        </div>
      )}
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" disabled={!valid} onClick={apply}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
