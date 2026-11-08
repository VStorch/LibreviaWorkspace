import { useEffect } from 'react'
import { create } from 'zustand'
import { MenuCommand } from '@shared/types.js'
import { clampZoom, createWheelZoom, zoomIn, zoomOut, type ZoomStep } from '@services/document/zoom.js'
import { currentPreferences, setPreference, usePreferences } from './preferences.js'

/** Not a preference: it depends on the window. The editor measures it. */
export const useFittedZoom = create<{ fitted: number }>(() => ({ fitted: 100 }))

export function setFittedZoom(fitted: number): void {
  if (useFittedZoom.getState().fitted !== fitted) useFittedZoom.setState({ fitted })
}

export function useEffectiveZoom(): number {
  const zoom = usePreferences((state) => state.preferences.zoom)
  const fit = usePreferences((state) => state.preferences.zoomFit)
  const fitted = useFittedZoom((state) => state.fitted)
  return fit ? fitted : zoom
}

let queued: Promise<void> = Promise.resolve()

/**
 * In a row: the preference changes only when main answers, and two wheel notches in the meantime
 * would both step from the same value.
 */
export function runZoomCommand(command: MenuCommand): Promise<void> {
  const run = queued.then(() => applyZoomCommand(command))
  queued = run.catch(() => undefined)
  return run
}

export function zoomCommandOf(step: ZoomStep): MenuCommand {
  return step === 'in' ? MenuCommand.ZoomIn : MenuCommand.ZoomOut
}

/** Not passive: Ctrl+wheel must not also scroll the sheets. */
export function useCtrlWheelZoom(target: { readonly current: HTMLElement | null }, mounted: boolean): void {
  useEffect(() => {
    const element = target.current
    if (element === null) return undefined
    const wheel = createWheelZoom()
    const onWheel = (event: WheelEvent): void => {
      if (!event.ctrlKey) return
      event.preventDefault()
      const step = wheel(event.deltaY, event.timeStamp)
      if (step !== null) void runZoomCommand(zoomCommandOf(step))
    }
    element.addEventListener('wheel', onWheel, { passive: false })
    return () => element.removeEventListener('wheel', onWheel)
  }, [target, mounted])
}

async function applyZoomCommand(command: MenuCommand): Promise<void> {
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
