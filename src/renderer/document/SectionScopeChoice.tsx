import { useT } from '../i18n.js'

export type SectionScope = 'section' | 'document'

/** "Nesta seção" ou "no documento todo", como no Word. */
export function SectionScopeChoice({
  name,
  scope,
  onChange,
}: {
  /** O grupo dos botões de rádio. */
  readonly name: string
  readonly scope: SectionScope
  readonly onChange: (scope: SectionScope) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__row">
      {(['section', 'document'] as const).map((choice) => (
        <label key={choice} className="popover__check">
          <input type="radio" name={name} checked={scope === choice} onChange={() => onChange(choice)} />
          {t(choice === 'section' ? 'document.pageSetup.thisSection' : 'document.pageSetup.wholeDocument')}
        </label>
      ))}
    </div>
  )
}
