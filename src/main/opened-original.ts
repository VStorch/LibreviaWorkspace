/**
 * The packages the open files came from, as they were on disk when opened, not as they are now,
 * which someone else may have changed. The bytes live **here**, not in the stateless sidecar: its
 * death does not cost the surgical save. One original per path, and of one format: a block's `oid`
 * only means something in the package it was read from. The renderer releases a path when it closes
 * that file.
 */

import { readFile } from 'node:fs/promises'
import { pathKey } from './fs/paths.js'

export const OriginalFormat = { Docx: 'docx', Odt: 'odt', Xlsx: 'xlsx' } as const
export type OriginalFormat = (typeof OriginalFormat)[keyof typeof OriginalFormat]

interface OpenedOriginal {
  readonly bytes: Buffer
  readonly format: OriginalFormat
  /** The minimal package of a document that had no original of this format: rebuilt on each save. */
  readonly created: boolean
}

const opened = new Map<string, OpenedOriginal>()

export function rememberOriginal(path: string, bytes: Buffer, format: OriginalFormat, created = false): void {
  opened.set(pathKey(path), { bytes, format, created })
}

export function releaseOriginal(path: string): void {
  opened.delete(pathKey(path))
}

export function resetOriginals(): void {
  opened.clear()
}

function openedFor(origin: string | null): OpenedOriginal | null {
  return origin === null ? null : (opened.get(pathKey(origin)) ?? null)
}

/** The format of the original kept for `origin`, or `null` when there is none. */
export function originalFormatOf(origin: string | null): OriginalFormat | null {
  return openedFor(origin)?.format ?? null
}

/** Only when the document comes from that very file, in that format. */
export function originalOf(origin: string | null, format: OriginalFormat): Buffer | null {
  const kept = openedFor(origin)
  return kept?.format === format ? kept.bytes : null
}

/** Whether the original kept for `origin` is a minimal package, not a file the user opened. */
export function isCreatedOriginal(origin: string | null): boolean {
  return openedFor(origin)?.created === true
}

/**
 * After a crash the original bytes died with the process: without rereading them, saving would be
 * refused.
 */
export async function adoptOriginal(path: string, format: OriginalFormat): Promise<boolean> {
  try {
    rememberOriginal(path, await readFile(path), format)
    return true
  } catch {
    releaseOriginal(path)
    return false
  }
}

export interface SavedOriginal {
  /** The starting package of the save, not the saved bytes; see `saveDocx`. */
  readonly original: Buffer
  readonly format: OriginalFormat
  readonly created?: boolean
}

/**
 * Called **after** the save reached the disk, so a failure does not leave the original pointing to
 * a path the document does not have. Applies to any destination: a `.docx` or `.odt` saved as
 * `.sdoc` keeps the same `oid`s.
 */
export function followOriginal(
  origin: string | null,
  destination: string,
  saved: SavedOriginal | null,
): void {
  const kept =
    saved === null
      ? openedFor(origin)
      : { bytes: saved.original, format: saved.format, created: saved.created === true }
  if (origin !== null) releaseOriginal(origin)
  if (kept !== null) rememberOriginal(destination, kept.bytes, kept.format, kept.created)
}
