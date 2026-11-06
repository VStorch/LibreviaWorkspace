/**
 * The package the open document came from, as it was on disk when opened, not as it is now, which
 * someone else may have changed. The bytes live **here**, not in the stateless sidecar: its death
 * does not cost the surgical save. One original at a time, and of one format: a block's `oid` only
 * means something in the package it was read from.
 */

import { readFile } from 'node:fs/promises'
import { normalizePath } from './fs/paths.js'

export const OriginalFormat = { Docx: 'docx', Odt: 'odt' } as const
export type OriginalFormat = (typeof OriginalFormat)[keyof typeof OriginalFormat]

interface OpenedOriginal {
  readonly path: string
  readonly bytes: Buffer
  readonly format: OriginalFormat
}

let opened: OpenedOriginal | null = null

/** Normalized, as `authorizePath` returns it and `origin` brings it. */
export function rememberOriginal(path: string, bytes: Buffer, format: OriginalFormat): void {
  opened = { path: normalizePath(path), bytes, format }
}

export function forgetOpenedOriginal(): void {
  opened = null
}

function openedFor(origin: string | null): OpenedOriginal | null {
  if (opened === null || origin === null || opened.path !== normalizePath(origin)) return null
  return opened
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

/**
 * After a crash the original bytes died with the process: without rereading them, saving would be
 * refused.
 */
export async function adoptOriginal(path: string, format: OriginalFormat): Promise<boolean> {
  try {
    rememberOriginal(path, await readFile(path), format)
    return true
  } catch {
    opened = null
    return false
  }
}

export interface SavedOriginal {
  /** The starting package of the save, not the saved bytes; see `saveDocx`. */
  readonly original: Buffer
  readonly format: OriginalFormat
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
  if (saved !== null) {
    rememberOriginal(destination, saved.original, saved.format)
    return
  }
  const kept = openedFor(origin)
  if (kept !== null) rememberOriginal(destination, kept.bytes, kept.format)
}
