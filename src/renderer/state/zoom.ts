import { create } from 'zustand'
import { MenuCommand } from '@shared/types.js'
import { clampZoom, zoomIn, zoomOut } from '@services/document/zoom.js'
import { currentPreferences, setPreference, usePreferences } from './preferences.js'

/**
 * O zoom que "ajustar à largura" dá na janela de agora.
 *
 * Não é preferência: depende do tamanho da janela, e gravá-lo faria a próxima
 * abertura herdar a largura de uma janela que já não existe. Quem mede é o
 * editor, que conhece a área de rolagem; a barra de status e os atalhos leem daqui.
 */
export const useFittedZoom = create<{ fitted: number }>(() => ({ fitted: 100 }))

export function setFittedZoom(fitted: number): void {
  if (useFittedZoom.getState().fitted !== fitted) useFittedZoom.setState({ fitted })
}

/** O zoom que se vê, em porcento. */
export function useEffectiveZoom(): number {
  const zoom = usePreferences((state) => state.preferences.zoom)
  const fit = usePreferences((state) => state.preferences.zoomFit)
  const fitted = useFittedZoom((state) => state.fitted)
  return fit ? fitted : zoom
}

/** Ampliar e reduzir partem do zoom que se vê — inclusive do ajustado à largura. */
export async function runZoomCommand(command: MenuCommand): Promise<void> {
  const preferences = currentPreferences()
  const now = preferences.zoomFit ? useFittedZoom.getState().fitted : preferences.zoom
  switch (command) {
    case MenuCommand.ZoomIn:
      return setPreference({ zoom: clampZoom(zoomIn(now)), zoomFit: false })
    case MenuCommand.ZoomOut:
      return setPreference({ zoom: clampZoom(zoomOut(now)), zoomFit: false })
    case MenuCommand.ZoomReset:
      return setPreference({ zoom: 100, zoomFit: false })
    case MenuCommand.ZoomFitWidth:
      return setPreference({ zoomFit: !preferences.zoomFit })
  }
}
