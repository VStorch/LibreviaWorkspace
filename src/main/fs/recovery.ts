/**
 * The draft **never** touches the user's file: overwriting it would turn "I didn't save" into "I
 * saved by accident". One draft per tab, named by its id. The folder comes as a parameter, to
 * test without Electron.
 */

import { randomUUID } from 'node:crypto'
import { mkdir, readdir, readFile, rename, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { draftIdSchema, MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { writeFileAtomic } from './atomic-write.js'
import { t } from '../i18n.js'
import { MAX_DRAFTS, MAX_FILE_NAME_LENGTH } from '@shared/limits.js'

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

/** Before drafts were kept per tab, there was this single file. */
const LEGACY_FILE_NAME = 'draft.json'
const DRAFTS_FOLDER = 'drafts'
const EXTENSION = '.json'

let folder: string | null = null

/** Once, when main starts. */
export function useRecoveryFolder(path: string): void {
  folder = join(path, 'recovery')
}

function directory(): string {
  if (folder === null) throw new Error(t('errors.recovery.folderNotConfigured'))
  return folder
}

function draftsDirectory(): string {
  return join(directory(), DRAFTS_FOLDER)
}

function file(draftId: string): string {
  return join(draftsDirectory(), `${draftIdSchema.parse(draftId)}${EXTENSION}`)
}

/** No `.bak`, since it is rewritten every few seconds; the swap is still atomic. */
export async function writeDraft(
  draftId: string,
  draft: Omit<Draft, 'savedAt' | 'draftId'>,
): Promise<number> {
  const target = file(draftId)
  const savedAt = Date.now()
  await mkdir(draftsDirectory(), { recursive: true })
  await writeFileAtomic(target, JSON.stringify({ ...draft, savedAt }), { backup: false })
  return savedAt
}

/** Unreadable means absent: an error here would be a second failure right after a crash. */
export async function readDraft(draftId: string): Promise<Draft | null> {
  const target = file(draftId)
  let text: string
  try {
    text = await readFile(target, 'utf8')
  } catch {
    return null
  }

  try {
    const parsed = draftSchema.safeParse(JSON.parse(text))
    return parsed.success ? { draftId, ...parsed.data } : null
  } catch {
    return null
  }
}

/** The most recent first; unreadable drafts are left out. */
export async function listDraftSummaries(): Promise<DraftSummary[]> {
  await migrateLegacyDraft()

  const drafts = await Promise.all((await draftIds()).map((draftId) => readDraft(draftId)))
  return drafts
    .filter((draft) => draft !== null)
    .map(({ draftId, path, name, kind, savedAt }) => ({ draftId, path, name, kind, savedAt }))
    .sort((a, b) => b.savedAt - a.savedAt)
    .slice(0, MAX_DRAFTS)
}

async function draftIds(): Promise<string[]> {
  const names = await readdir(draftsDirectory()).catch(() => [])
  return names
    .filter((name) => name.endsWith(EXTENSION))
    .map((name) => name.slice(0, -EXTENSION.length))
    .filter((draftId) => draftIdSchema.safeParse(draftId).success)
}

/** Moved as it is, even if corrupt: reading it decides, as for any other draft. */
async function migrateLegacyDraft(): Promise<void> {
  const legacy = join(directory(), LEGACY_FILE_NAME)
  try {
    await mkdir(draftsDirectory(), { recursive: true })
    await rename(legacy, file(randomUUID()))
  } catch {
    // Nothing to migrate is the normal case.
  }
}

/** Nothing to delete is the normal case. */
export async function discardDraft(draftId: string): Promise<void> {
  const target = file(draftId)
  await rm(target, { force: true }).catch(() => undefined)
}
