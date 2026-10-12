import { z } from 'zod'
import { MAX_ZOOM, MIN_ZOOM } from '@services/document/zoom.js'
import { PDF_VIEWER_ORIGIN } from './constants.js'

/**
 * What the app and the PDF viewer frame say to each other through `postMessage`. Each side reads
 * the other's messages as untrusted: the viewer runs code that parses the PDF.
 */

const MAX_QUERY_LENGTH = 1_000
const MAX_URL_LENGTH = 8_192
const MAX_PAGES = 1_000_000
/** PDF 2.0 cuts a password at 127 bytes; this only keeps a message from growing without end. */
const MAX_PASSWORD_LENGTH = 1_024

const zoomSchema = z.number().int().min(MIN_ZOOM).max(MAX_ZOOM)
const pageSchema = z.number().int().min(1).max(MAX_PAGES)
const countSchema = z.number().int().min(0).max(MAX_PAGES)

export const PAGE_WIDTH = 'page-width'

export const viewerCommandSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('open'),
    url: z.string().startsWith(`${PDF_VIEWER_ORIGIN}/document/`).max(MAX_URL_LENGTH),
  }),
  z.object({ type: z.literal('theme'), theme: z.enum(['light', 'dark']) }),
  z.object({ type: z.literal('zoom'), zoom: z.union([zoomSchema, z.literal(PAGE_WIDTH)]) }),
  z.object({
    type: z.literal('find'),
    query: z.string().max(MAX_QUERY_LENGTH),
    again: z.boolean(),
    backwards: z.boolean(),
    caseSensitive: z.boolean(),
  }),
  z.object({ type: z.literal('find-close') }),
  z.object({ type: z.literal('password'), password: z.string().max(MAX_PASSWORD_LENGTH) }),
])

export const viewerEventSchema = z.discriminatedUnion('type', [
  z.object({ type: z.literal('ready') }),
  z.object({ type: z.literal('loaded'), pages: countSchema }),
  z.object({ type: z.literal('page'), page: pageSchema }),
  z.object({ type: z.literal('zoom'), zoom: zoomSchema }),
  z.object({ type: z.literal('zoom-in') }),
  z.object({ type: z.literal('found'), current: countSchema, total: countSchema }),
  /** `retry`: the last password was wrong. */
  z.object({ type: z.literal('password'), retry: z.boolean() }),
  z.object({ type: z.literal('link'), url: z.string().max(MAX_URL_LENGTH) }),
  z.object({ type: z.literal('error'), reason: z.enum(['password', 'invalid', 'unknown']) }),
])

export type ViewerCommand = z.infer<typeof viewerCommandSchema>
export type ViewerEvent = z.infer<typeof viewerEventSchema>
