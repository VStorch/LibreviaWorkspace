import { create } from 'zustand'
import type { ViewerCommand } from '@shared/pdf-viewer-protocol.js'
import { createTabSlot } from '../state/tabs.js'

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

/**
 * What the active tab's viewer reported, for the status bar and the find bar. Window-wide: only the
 * active tab's viewer is mounted, and connecting or leaving clears it.
 */
export const usePdfViewer = create<PdfViewerState>(() => NOTHING_SHOWN)

const post = createTabSlot<(command: ViewerCommand) => void>()

/** Commands go to the active tab's viewer. Returns the disconnect function. */
export function connectViewer(tabId: string, send: (command: ViewerCommand) => void): () => void {
  const release = post.hold(tabId, send)
  usePdfViewer.setState(NOTHING_SHOWN)
  return () => {
    release()
    usePdfViewer.setState(NOTHING_SHOWN)
  }
}

export function sendToViewer(command: ViewerCommand): void {
  post.active()?.(command)
}
