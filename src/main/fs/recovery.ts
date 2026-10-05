/**
 * The draft **never** touches the user's file: overwriting it would turn "I didn't save" into "I
 * saved by accident". One draft at a time, like one file at a time. The folder comes as a
 * parameter, to test without Electron.
 */

import { mkdir, readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { writeFileAtomic } from './atomic-write.js'
import { t } from '../i18n.js'
import { MAX_FILE_NAME_LENGTH } from '@shared/limits.js'

/** The summary, without content, lives in `shared/types.ts`. */
export interface Draft extends DraftSummary {
  readonly content: string
}

const draftSchema = z.object({
  path: z.string().min(1).nullable(),
  name: z.string().min(1).max(MAX_FILE_NAME_LENGTH),
  kind: z.enum([DocumentKind.Document, DocumentKind.Spreadsheet]),
  content: z.string().max(MAX_TEXT_LENGTH),
  savedAt: z.number().int().positive(),
})

const FILE_NAME = 'rascunho.json'

let folder: string | null = null

/** Once, when main starts. */
export function useRecoveryFolder(path: string): void {
  folder = join(path, 'recuperacao')
}

function directory(): string {
  if (folder === null) throw new Error(t('errors.recovery.folderNotConfigured'))
  return folder
}

function file(): string {
  return join(directory(), FILE_NAME)
}

/** No `.bak`, since it is rewritten every few seconds; the swap is still atomic. */
export async function writeDraft(draft: Omit<Draft, 'savedAt'>): Promise<number> {
  const savedAt = Date.now()
  await mkdir(directory(), { recursive: true })
  await writeFileAtomic(file(), JSON.stringify({ ...draft, savedAt }), { backup: false })
  return savedAt
}

/** Unreadable means absent: an error here would be a second failure right after a crash. */
export async function readDraft(): Promise<Draft | null> {
  let text: string
  try {
    text = await readFile(file(), 'utf8')
  } catch {
    return null
  }

  try {
    const parsed = draftSchema.safeParse(JSON.parse(text))
    return parsed.success ? parsed.data : null
  } catch {
    return null
  }
}

export async function readDraftSummary(): Promise<DraftSummary | null> {
  const draft = await readDraft()
  if (draft === null) return null

  return { path: draft.path, name: draft.name, kind: draft.kind, savedAt: draft.savedAt }
}

/** Nothing to delete is the normal case. */
export async function discardDraft(): Promise<void> {
  await rm(file(), { force: true }).catch(() => undefined)
}
