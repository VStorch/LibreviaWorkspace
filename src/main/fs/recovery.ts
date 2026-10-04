/**
 * O rascunho **nunca** toca o arquivo do usuário: gravar por cima transformaria
 * "não salvei" em "salvei sem querer". Um rascunho por vez, como um arquivo por
 * vez. A pasta chega por parâmetro, para testar sem Electron.
 */

import { mkdir, readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { writeFileAtomic } from './atomic-write.js'
import { t } from '../i18n.js'
import { MAX_FILE_NAME_LENGTH } from '@shared/limits.js'

/** O resumo, sem conteúdo, mora em `shared/types.ts`. */
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

/** Uma vez, na subida do main. */
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

/** Sem `.bak`, porque é reescrito a cada poucos segundos; a troca continua atômica. */
export async function writeDraft(draft: Omit<Draft, 'savedAt'>): Promise<number> {
  const savedAt = Date.now()
  await mkdir(directory(), { recursive: true })
  await writeFileAtomic(file(), JSON.stringify({ ...draft, savedAt }), { backup: false })
  return savedAt
}

/** Ilegível é ausente: um erro aqui seria uma segunda falha logo depois de uma queda. */
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

/** Não haver o que apagar é o caso normal. */
export async function discardDraft(): Promise<void> {
  await rm(file(), { force: true }).catch(() => undefined)
}
