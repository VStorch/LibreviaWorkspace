import { getDocument, type PDFDocumentLoadingTask } from 'pdfjs-dist'
import { EventBus, PDFFindController, PDFLinkService, PDFViewer } from 'pdfjs-dist/web/pdf_viewer.mjs'
import { clampZoom } from '@services/document/zoom.js'
import { PDF_VIEWER_ORIGIN } from '@shared/constants.js'
import { PAGE_WIDTH, type ViewerCommand } from '@shared/pdf-viewer-protocol.js'
import { isPdfZoomInKey } from '@shared/shortcuts.js'
import { tellApp } from './channel.js'
import { rangedSource, type PdfSource } from './ranged-source.js'

type Zoom = Extract<ViewerCommand, { type: 'zoom' }>['zoom']

const RESOURCES = `${PDF_VIEWER_ORIGIN}/pdfjs`

/** pdf.js names its failures; the app shows its own message for each. */
const FAILURES: Readonly<Record<string, 'password' | 'invalid'>> = {
  PasswordException: 'password',
  InvalidPDFException: 'invalid',
}

/**
 * pdf.js draws the pages; everything around them (page count, zoom, search) belongs to the app.
 * Without a scripting manager the PDF's JavaScript never runs, and XFA forms stay off.
 */
export function createViewer(container: HTMLDivElement): (command: ViewerCommand) => void {
  const eventBus = new EventBus()
  const linkService = new PDFLinkService({ eventBus })
  const findController = new PDFFindController({ eventBus, linkService })
  const viewer = new PDFViewer({ container, eventBus, linkService, findController, removePageBorders: true })
  linkService.setViewer(viewer)

  let zoom: Zoom = PAGE_WIDTH
  let task: PDFDocumentLoadingTask | null = null

  reportToApp(eventBus)
  eventBus.on('pagesinit', () => applyZoom(viewer, zoom))
  sendLinksToApp(container)
  sendZoomKeysToApp()

  return (command) => {
    switch (command.type) {
      case 'open':
        void task?.destroy()
        task = null
        rangedSource(command.url, reportFailure).then(
          (source) => (task = open(source, viewer, linkService)),
          reportFailure,
        )
        return
      case 'theme':
        document.documentElement.dataset['theme'] = command.theme
        return
      case 'zoom':
        zoom = command.zoom
        if (viewer.pdfDocument !== null) applyZoom(viewer, zoom)
        return
      case 'find':
        eventBus.dispatch('find', {
          source: null,
          type: command.again ? 'again' : '',
          query: command.query,
          caseSensitive: command.caseSensitive,
          entireWord: false,
          highlightAll: true,
          findPrevious: command.backwards,
          matchDiacritics: false,
        })
        return
      case 'find-close':
        eventBus.dispatch('findbarclose', { source: null })
        return
    }
  }
}

function open(source: PdfSource, viewer: PDFViewer, linkService: PDFLinkService): PDFDocumentLoadingTask {
  const task = getDocument({
    ...source,
    // Only the pieces the pages on screen need, and the text a search reads.
    disableAutoFetch: true,
    enableXfa: false,
    cMapUrl: `${RESOURCES}/cmaps/`,
    standardFontDataUrl: `${RESOURCES}/standard_fonts/`,
    wasmUrl: `${RESOURCES}/wasm/`,
    iccUrl: `${RESOURCES}/iccs/`,
  })

  task.promise.then(
    (pdf) => {
      viewer.setDocument(pdf)
      linkService.setDocument(pdf)
      tellApp({ type: 'loaded', pages: pdf.numPages })
    },
    (error: unknown) => {
      // Destroyed because another PDF took its place: nothing to report.
      if (task.destroyed) return
      const name = error instanceof Error ? error.name : ''
      tellApp({ type: 'error', reason: FAILURES[name] ?? 'unknown' })
    },
  )
  return task
}

function applyZoom(viewer: PDFViewer, zoom: Zoom): void {
  if (zoom === PAGE_WIDTH) viewer.currentScaleValue = PAGE_WIDTH
  else viewer.currentScale = zoom / 100
}

function reportFailure(): void {
  tellApp({ type: 'error', reason: 'unknown' })
}

interface FindCount {
  readonly current: number
  readonly total: number
}

function reportToApp(eventBus: EventBus): void {
  eventBus.on('pagechanging', ({ pageNumber }: { pageNumber: number }) => {
    tellApp({ type: 'page', page: pageNumber })
  })
  eventBus.on('scalechanging', ({ scale }: { scale: number }) => {
    tellApp({ type: 'zoom', zoom: clampZoom(scale * 100) })
  })

  const reportFound = ({ matchesCount }: { matchesCount?: FindCount }): void => {
    tellApp({ type: 'found', current: matchesCount?.current ?? 0, total: matchesCount?.total ?? 0 })
  }
  eventBus.on('updatefindmatchescount', reportFound)
  eventBus.on('updatefindcontrolstate', reportFound)
}

/** A link out of the PDF goes to the app, which asks before opening the browser. */
function sendLinksToApp(container: HTMLElement): void {
  container.addEventListener('click', (event) => {
    const anchor = event.target instanceof Element ? event.target.closest('a') : null
    if (anchor === null || anchor.href === '' || anchor.href.startsWith(PDF_VIEWER_ORIGIN)) return
    event.preventDefault()
    tellApp({ type: 'link', url: anchor.href })
  })
}

/** The zoom is the app's, shared with the editor: the frame only asks for a step. */
function sendZoomKeysToApp(): void {
  window.addEventListener('keydown', (event) => {
    if (!isPdfZoomInKey(event)) return
    event.preventDefault()
    tellApp({ type: 'zoom-in' })
  })
}
