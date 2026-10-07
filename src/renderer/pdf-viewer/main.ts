import * as pdfjs from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import 'pdfjs-dist/web/pdf_viewer.css'
import '../tokens.css'
import './viewer.css'
import { listenToApp, tellApp } from './channel.js'

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl
// The viewer components read the library from the global scope, so it goes there before they load.
Object.assign(globalThis, { pdfjsLib: pdfjs })
const { createViewer } = await import('./viewer.js')

const container = document.getElementById('viewerContainer')
if (container instanceof HTMLDivElement) {
  listenToApp(createViewer(container))
  tellApp({ type: 'ready' })
}
