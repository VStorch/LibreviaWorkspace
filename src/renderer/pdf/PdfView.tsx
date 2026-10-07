import { useEffect, useRef } from 'react'
import { useT } from '../i18n.js'

interface ViewBounds {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/**
 * Any dialog the app draws. Main places the viewer above the page, so it would cover the dialog:
 * while one is open, the viewer hides. A future dialog only needs `aria-modal`.
 */
const MODAL_SELECTOR = '[aria-modal="true"]'

/** Reserves the area; main draws Chromium's viewer over it. */
export function PdfView({ path }: { readonly path: string }): React.JSX.Element {
  const t = useT()
  const area = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const element = area.current
    if (element === null) return undefined

    let last = JSON.stringify(boundsOf(element))
    void window.api.pdf.attach({ path, bounds: boundsOf(element) })

    const sync = (): void => {
      const bounds = document.querySelector(MODAL_SELECTOR) === null ? boundsOf(element) : null
      const key = JSON.stringify(bounds)
      if (key === last) return
      last = key
      void window.api.pdf.layout({ bounds })
    }

    const resize = new ResizeObserver(sync)
    resize.observe(element)
    const modals = new MutationObserver(sync)
    modals.observe(document.body, { childList: true, subtree: true, attributeFilter: ['aria-modal'] })
    window.addEventListener('resize', sync)
    sync()

    return () => {
      resize.disconnect()
      modals.disconnect()
      window.removeEventListener('resize', sync)
      void window.api.pdf.detach({})
    }
  }, [path])

  return <div ref={area} className="pdf-view" role="region" aria-label={t('shell.pdf.viewer')} />
}

function boundsOf(element: HTMLElement): ViewBounds {
  const rect = element.getBoundingClientRect()
  return {
    x: Math.max(0, Math.round(rect.left)),
    y: Math.max(0, Math.round(rect.top)),
    width: Math.max(0, Math.round(rect.width)),
    height: Math.max(0, Math.round(rect.height)),
  }
}
