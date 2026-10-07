import { pageDimensionsMm } from '@services/document/model.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentKind, MenuCommand } from '@shared/types.js'
import { MAX_ZOOM, MIN_ZOOM } from '@services/document/zoom.js'
import { usePreferences } from '../state/preferences.js'
import { runZoomCommand, useEffectiveZoom } from '../state/zoom.js'
import { usePdfViewer } from '../pdf/viewer-state.js'

export function StatusBar(): React.JSX.Element {
  const t = useT()
  const state = useWorkspace()
  const { width, height } = pageDimensionsMm(state.page)

  // A spreadsheet has no pages, words or characters: showing them there would be false information in the
  // place of the true one.
  const sheet = state.workbook?.sheets[state.workbook.activeSheet]

  return (
    <footer className="statusbar">
      {/* `rtl` cuts the start of a long path, keeping the name; `bdi` keeps its slashes in place. */}
      <span className="statusbar__file" title={state.file?.path ?? undefined}>
        <bdi dir="ltr">{state.file?.path ?? t('shell.statusBar.notSavedYet')}</bdi>
      </span>

      <span className="statusbar__spacer" />

      {state.file?.kind === DocumentKind.Pdf ? (
        <PdfMetrics />
      ) : sheet === undefined ? (
        <>
          <span className="statusbar__metric">
            {state.page.size} {width} × {height} mm
          </span>
          <span className="statusbar__metric">{t('shell.statusBar.pages', { count: state.pageCount })}</span>
          <span className="statusbar__metric">
            {t('shell.statusBar.words', { count: state.stats.words })}
          </span>
          <span className="statusbar__metric">
            {t('shell.statusBar.characters', { count: state.stats.characters })}
          </span>
          {state.trackChanges === true && (
            <span
              className="statusbar__metric"
              data-testid="track-changes-status"
              title={t('revisions.trackHint')}
            >
              {t('revisions.trackOn')}
            </span>
          )}
          <ZoomControls />
        </>
      ) : (
        <>
          <span className="statusbar__metric">{sheet.name}</span>
          <span className="statusbar__metric">
            {t('shell.statusBar.filledCells', { count: Object.keys(sheet.cells).length })}
          </span>
        </>
      )}

      <span className={state.isDirty ? 'statusbar__state statusbar__state--dirty' : 'statusbar__state'}>
        {state.busy
          ? t('shell.statusBar.working')
          : state.isDirty
            ? t('shell.statusBar.unsaved')
            : t('shell.statusBar.saved')}
      </span>
    </footer>
  )
}

function PdfMetrics(): React.JSX.Element {
  const t = useT()
  const page = usePdfViewer((viewer) => viewer.page)
  const pages = usePdfViewer((viewer) => viewer.pages)
  return (
    <>
      {pages > 0 && (
        <span className="statusbar__metric" data-testid="pdf-page">
          {t('shell.statusBar.pageOf', { page, pages })}
        </span>
      )}
      <ZoomControls />
    </>
  )
}

function ZoomControls(): React.JSX.Element {
  const t = useT()
  const zoom = useEffectiveZoom()
  const zoomFit = usePreferences((preferences) => preferences.preferences.zoomFit)
  return (
    <span className="statusbar__zoom" role="group" aria-label={t('shell.statusBar.zoomLevel')}>
      <button
        type="button"
        className="statusbar__zoom-button"
        aria-label={t('shell.statusBar.zoomOut')}
        title={t('shell.statusBar.zoomOut')}
        disabled={zoom <= MIN_ZOOM}
        onClick={() => void runZoomCommand(MenuCommand.ZoomOut)}
      >
        −
      </button>
      <button
        type="button"
        className="statusbar__zoom-level"
        title={t('shell.statusBar.zoomLevel')}
        onClick={() => void runZoomCommand(MenuCommand.ZoomReset)}
      >
        {zoom}%
      </button>
      <button
        type="button"
        className="statusbar__zoom-button"
        aria-label={t('shell.statusBar.zoomIn')}
        title={t('shell.statusBar.zoomIn')}
        disabled={zoom >= MAX_ZOOM}
        onClick={() => void runZoomCommand(MenuCommand.ZoomIn)}
      >
        +
      </button>
      <button
        type="button"
        className={`statusbar__zoom-button${zoomFit ? ' statusbar__zoom-button--active' : ''}`}
        aria-pressed={zoomFit}
        aria-label={t('shell.statusBar.zoomFitWidth')}
        title={t('shell.statusBar.zoomFitWidth')}
        onClick={() => void runZoomCommand(MenuCommand.ZoomFitWidth)}
      >
        ↔
      </button>
    </span>
  )
}
