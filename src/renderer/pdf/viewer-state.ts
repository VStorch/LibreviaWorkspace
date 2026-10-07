import { create } from 'zustand'
import type { ViewerCommand } from '@shared/pdf-viewer-protocol.js'

export interface FindCount {
  readonly current: number
  readonly total: number
}

interface PdfViewerState {
  readonly page: number
  readonly pages: number
  /** `null` before any search. */
  readonly found: FindCount | null
}

const NOTHING_SHOWN: PdfViewerState = { page: 0, pages: 0, found: null }

/** What the open PDF's viewer reported, for the status bar and the find bar. */
export const usePdfViewer = create<PdfViewerState>(() => NOTHING_SHOWN)

let post: ((command: ViewerCommand) => void) | null = null

/** One viewer at a time: the open PDF's. Returns the disconnect function. */
export function connectViewer(send: (command: ViewerCommand) => void): () => void {
  post = send
  usePdfViewer.setState(NOTHING_SHOWN)
  return () => {
    if (post === send) post = null
    usePdfViewer.setState(NOTHING_SHOWN)
  }
}

export function sendToViewer(command: ViewerCommand): void {
  post?.(command)
}
