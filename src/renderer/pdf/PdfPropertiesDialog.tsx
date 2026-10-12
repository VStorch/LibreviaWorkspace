import { useEffect, useRef, useState } from 'react'
import { propertiesOf } from '@services/file/properties.js'
import type { IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import {
  FilePropertiesDialog,
  type EditableKey,
  type StatisticRow,
} from '../components/FilePropertiesDialog.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { usePdfViewer } from './viewer-state.js'

type Inspection = IpcResponse<typeof IpcChannel.PdfProperties>

/** What the information dictionary and XMP share. */
const PDF_FIELDS: readonly EditableKey[] = ['title', 'subject', 'creator', 'keywords']

/**
 * The sidecar reads the file when the dialog opens. The properties go to the workspace only the
 * first time: after that they are the ones being edited.
 */
export function PdfPropertiesDialog({
  path,
  onClose,
}: {
  readonly path: string
  readonly onClose: () => void
}): React.JSX.Element | null {
  const t = useT()
  const inspection = useInspection(path, onClose)
  if (inspection === null) return null

  if (inspection.encrypted) {
    return (
      <div
        className="popover"
        role="dialog"
        aria-label={t('shell.pdf.propertiesTitle')}
        onKeyDown={(event) => {
          if (event.key === 'Escape') onClose()
        }}
      >
        <p className="popover__hint">{t('shell.pdf.encryptedProperties')}</p>
        <div className="popover__actions">
          <span className="popover__spacer" />
          <button type="button" className="btn btn--primary" autoFocus onClick={onClose}>
            {t('document.common.close')}
          </button>
        </div>
      </div>
    )
  }

  return <EditablePdfProperties inspection={inspection} onClose={onClose} />
}

function EditablePdfProperties({
  inspection,
  onClose,
}: {
  readonly inspection: Inspection
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const pages = usePdfViewer((viewer) => viewer.pages)

  const statistics: readonly StatisticRow[] = [
    ['shell.pdf.application', inspection.application ?? '—'],
    ['shell.pdf.producer', inspection.producer ?? '—'],
    ['shell.pdf.version', inspection.version],
    ['document.wordCount.pages', pages.toLocaleString(language)],
  ]

  return (
    <FilePropertiesDialog
      title="shell.pdf.propertiesTitle"
      statistics={statistics}
      fields={PDF_FIELDS}
      authorship={false}
      notice={inspection.signed ? t('shell.pdf.signedWarning') : undefined}
      onClose={onClose}
    />
  )
}

function useInspection(path: string, onClose: () => void): Inspection | null {
  const [inspection, setInspection] = useState<Inspection | null>(null)
  // The latest, without reading the file again when the parent draws a new function.
  const close = useRef(onClose)
  close.current = onClose

  useEffect(() => {
    let current = true
    void (async () => {
      const result = await window.api.pdf.properties({ path })
      if (!current) return
      if (!result.ok) {
        useWorkspace.getState().showError(result.error)
        close.current()
        return
      }
      if (useWorkspace.getState().properties === undefined) {
        useWorkspace.setState({ properties: propertiesOf(result.data.properties ?? {}) })
      }
      setInspection(result.data)
    })()
    return () => {
      current = false
    }
  }, [path])

  return inspection
}
