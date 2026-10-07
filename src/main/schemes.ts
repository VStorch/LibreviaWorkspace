import { protocol } from 'electron'
import { PDF_VIEWER_SCHEME } from '@shared/constants.js'
import { FONT_SCHEME } from './fonts.js'

/**
 * A real origin for each scheme: without one, Chromium refuses fonts through CORS, and the PDF
 * viewer could neither fetch the file nor start its worker.
 */
const ORIGIN_PRIVILEGES = { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } as const

/** Before `app.whenReady()`, and only once: Chromium takes a single list. */
export function registerAppSchemes(): void {
  protocol.registerSchemesAsPrivileged(
    [FONT_SCHEME, PDF_VIEWER_SCHEME].map((scheme) => ({ scheme, privileges: ORIGIN_PRIVILEGES })),
  )
}
