import { useEffect, useState } from 'react'
import { PrivacyCategory, foundCategories, type PrivacyReport } from '@services/file/privacy.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { useT } from '../i18n.js'
import { anonymousName, inspectionNeedsSave } from '../state/privacy-actions.js'
import { useWorkspace } from '../state/workspace.js'

type Phase =
  | { readonly step: 'save' }
  | { readonly step: 'loading' }
  | { readonly step: 'report'; readonly report: PrivacyReport }
  | { readonly step: 'removed' }

const LABELS: Record<PrivacyCategory, MessageKey> = {
  [PrivacyCategory.Names]: 'privacy.names',
  [PrivacyCategory.Paths]: 'privacy.paths',
  [PrivacyCategory.CustomProperties]: 'privacy.customProperties',
  [PrivacyCategory.Thumbnail]: 'privacy.thumbnail',
  [PrivacyCategory.ImageMetadata]: 'privacy.imageMetadata',
}

/** Long lists would push the buttons off the screen. */
const MAX_LISTED = 8

/**
 * File → Inspect document: what the file reveals besides its content, as Word's inspector shows it.
 * Removing marks the file modified; only saving writes it.
 */
export function PrivacyInspector(): React.JSX.Element {
  const t = useT()
  const [phase, setPhase] = useState<Phase>(() =>
    inspectionNeedsSave(useWorkspace.getState()) ? { step: 'save' } : { step: 'loading' },
  )
  const close = (): void => useWorkspace.getState().setInspector(false)

  useEffect(() => {
    if (phase.step !== 'loading') return
    let current = true
    void useWorkspace
      .getState()
      .inspectPersonalData()
      .then((report) => {
        if (!current) return
        if (report === null) close()
        else setPhase({ step: 'report', report })
      })
    return () => {
      current = false
    }
  }, [phase.step])

  async function saveFirst(): Promise<void> {
    if (await useWorkspace.getState().save()) setPhase({ step: 'loading' })
  }

  return (
    <div
      className="popover popover--wide properties"
      role="dialog"
      aria-label={t('privacy.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') close()
      }}
    >
      {phase.step === 'save' && (
        <Message text={t('privacy.needsSave')} onClose={close}>
          <button type="button" className="btn btn--primary" autoFocus onClick={() => void saveFirst()}>
            {t('privacy.saveAndInspect')}
          </button>
        </Message>
      )}
      {phase.step === 'loading' && <p className="popover__hint">{t('privacy.loading')}</p>}
      {phase.step === 'report' && (
        <Report report={phase.report} onRemoved={() => setPhase({ step: 'removed' })} onClose={close} />
      )}
      {phase.step === 'removed' && <Message text={t('privacy.removed')} onClose={close} />}
    </div>
  )
}

function Message({
  text,
  onClose,
  children,
}: {
  readonly text: string
  readonly onClose: () => void
  readonly children?: React.ReactNode
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      <p className="popover__hint" role="status">
        {text}
      </p>
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" autoFocus={children === undefined} onClick={onClose}>
          {t('document.common.close')}
        </button>
        {children}
      </div>
    </>
  )
}

function Report({
  report,
  onRemoved,
  onClose,
}: {
  readonly report: PrivacyReport
  readonly onRemoved: () => void
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const readOnly = useWorkspace((state) => state.readOnly)
  const found = foundCategories(report)
  const [chosen, setChosen] = useState<ReadonlySet<PrivacyCategory>>(() => new Set(found))
  const [removing, setRemoving] = useState(false)

  async function remove(): Promise<void> {
    setRemoving(true)
    if (await useWorkspace.getState().removePersonalData([...chosen])) onRemoved()
    else setRemoving(false)
  }

  const toggle = (category: PrivacyCategory, checked: boolean): void => {
    const next = new Set(chosen)
    if (checked) next.add(category)
    else next.delete(category)
    setChosen(next)
  }

  return (
    <>
      <p className="popover__hint">{t(found.length === 0 ? 'privacy.nothingFound' : 'privacy.intro')}</p>
      {found.map((category, index) => (
        <div key={category}>
          <label className="popover__check">
            <input
              type="checkbox"
              checked={chosen.has(category)}
              autoFocus={index === 0}
              disabled={readOnly}
              onChange={(event) => toggle(category, event.target.checked)}
            />
            <strong>{t(LABELS[category])}</strong>
          </label>
          <p className="popover__hint privacy__detail">{detailOf(category, report, t)}</p>
        </div>
      ))}
      {report.removePersonalInformation && (
        <p className="popover__hint" role="note">
          {t('privacy.alreadyAnonymous', { anonymous: anonymousName() })}
        </p>
      )}
      {readOnly && found.length > 0 && <p className="popover__error">{t('privacy.readOnly')}</p>}
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" autoFocus={found.length === 0} onClick={onClose}>
          {t('document.common.close')}
        </button>
        {found.length > 0 && (
          <button
            type="button"
            className="btn btn--primary"
            disabled={readOnly || removing || chosen.size === 0}
            onClick={() => void remove()}
          >
            {t('privacy.remove')}
          </button>
        )}
      </div>
    </>
  )
}

function listed(items: readonly string[]): string {
  const shown = items.slice(0, MAX_LISTED).join(', ')
  return items.length > MAX_LISTED ? `${shown}… (+${items.length - MAX_LISTED})` : shown
}

function detailOf(category: PrivacyCategory, report: PrivacyReport, t: ReturnType<typeof useT>): string {
  switch (category) {
    case PrivacyCategory.Names:
      return t('privacy.namesDetail', { anonymous: anonymousName(), names: listed(report.names) })
    case PrivacyCategory.Paths:
      return t('privacy.pathsDetail', { paths: listed(report.paths) })
    case PrivacyCategory.CustomProperties:
      return t('privacy.customPropertiesDetail', { names: listed(report.customProperties) })
    case PrivacyCategory.Thumbnail:
      return t('privacy.thumbnailDetail')
    case PrivacyCategory.ImageMetadata:
      return t('privacy.imageMetadataDetail', { count: report.imagesWithMetadata })
  }
}
