import { pageDimensionsMm } from '@services/document/model.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { MenuCommand } from '@shared/types.js'
import { MAX_ZOOM, MIN_ZOOM } from '@services/document/zoom.js'
import { usePreferences } from '../state/preferences.js'
import { runZoomCommand, useEffectiveZoom } from '../state/zoom.js'

export function StatusBar(): React.JSX.Element {
  const t = useT()
  const state = useWorkspace()
  const { width, height } = pageDimensionsMm(state.page)

  // Planilha não tem página, palavra nem caractere: mostrar isso ali seria
  // informação falsa ocupando o lugar da verdadeira.
  const sheet = state.workbook?.sheets[state.workbook.activeSheet]
  const zoom = useEffectiveZoom()
  const zoomFit = usePreferences((preferences) => preferences.preferences.zoomFit)

  return (
    <footer className="statusbar">
      <span className="statusbar__file" title={state.file?.path ?? undefined}>
        {state.file?.path ?? t('shell.statusBar.notSavedYet')}
      </span>

      <span className="statusbar__spacer" />

      {sheet === undefined ? (
        <>
          <span className="statusbar__metric">
            {state.page.size} {width} × {height} mm
          </span>
          {/* Sem "≈": a tela agora pagina de verdade, e o número é o mesmo
              que a pessoa vê nas folhas. Prometer aproximação quando a conta
              está certa ensina a desconfiar de um número bom. */}
          <span className="statusbar__metric">
            {t('shell.statusBar.pages', { count: state.estimatedPages })}
          </span>
          <span className="statusbar__metric">
            {t('shell.statusBar.words', { count: state.stats.words })}
          </span>
          <span className="statusbar__metric">
            {t('shell.statusBar.characters', { count: state.stats.characters })}
          </span>
          {/* O zoom da folha, os mesmos comandos do menu Exibir. */}
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
        </>
      ) : (
        <>
          <span className="statusbar__metric">{sheet.name}</span>
          <span className="statusbar__metric">
            {t('shell.statusBar.filledCells', { count: Object.keys(sheet.cells).length })}
          </span>
        </>
      )}

      {/* Indicador de alterações não salvas: repete o marcador do título da
          janela, para que o estado seja legível sem sair do conteúdo. */}
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
