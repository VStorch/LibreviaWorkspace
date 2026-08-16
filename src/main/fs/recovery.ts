/**
 * Rascunho de recuperação: o que estava na tela quando o aplicativo caiu.
 *
 * Duas escolhas de desenho valem a pena registrar.
 *
 * **O rascunho nunca toca o arquivo do usuário.** Gravar por cima dele de tempos
 * em tempos transformaria "não salvei" em "salvei sem querer" — e desfazer isso
 * exigiria o `.bak`, que existe para outro problema. O rascunho vive na pasta do
 * aplicativo, e a decisão de escrever no arquivo continua sendo só do usuário.
 *
 * **Um rascunho por vez**, porque o aplicativo edita um arquivo por vez. Uma
 * fila de rascunhos precisaria de uma tela para escolher entre eles, e essa tela
 * só apareceria depois de uma queda — o pior momento para pedir uma decisão.
 *
 * Este módulo não importa `electron` de propósito: a pasta chega por parâmetro,
 * e é o que permite testá-lo sem subir um aplicativo inteiro.
 */

import { mkdir, readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { MAX_TEXT_LENGTH } from '@shared/ipc.js'
import { writeFileAtomic } from './atomic-write.js'

/** O rascunho inteiro. O resumo, sem conteúdo, mora em `shared/types.ts`. */
export interface Draft extends DraftSummary {
  readonly content: string
}

const draftSchema = z.object({
  path: z.string().min(1).nullable(),
  name: z.string().min(1).max(255),
  kind: z.enum([DocumentKind.Document, DocumentKind.Spreadsheet]),
  content: z.string().max(MAX_TEXT_LENGTH),
  savedAt: z.number().int().positive(),
})

const FILE_NAME = 'rascunho.json'

let folder: string | null = null

/** Aponta o rascunho para uma pasta. Chamado uma vez, na subida do main. */
export function useRecoveryFolder(path: string): void {
  folder = join(path, 'recuperacao')
}

function directory(): string {
  if (folder === null) throw new Error('a pasta de recuperação não foi configurada')
  return folder
}

function file(): string {
  return join(directory(), FILE_NAME)
}

/**
 * Grava o rascunho.
 *
 * Sem cópia `.bak`: o rascunho é reescrito a cada poucos segundos, e guardar a
 * versão anterior de cada uma delas só dobraria a escrita. A troca continua
 * atômica, que é o que impede um rascunho pela metade de sobreviver a uma queda.
 */
export async function writeDraft(draft: Omit<Draft, 'savedAt'>): Promise<number> {
  const savedAt = Date.now()
  await mkdir(directory(), { recursive: true })
  await writeFileAtomic(file(), JSON.stringify({ ...draft, savedAt }), { backup: false })
  return savedAt
}

/**
 * Lê o rascunho, ou `null` quando não há.
 *
 * Rascunho ilegível é tratado como ausente, e de propósito: ele existe para
 * salvar o dia depois de uma queda, e um erro de leitura dele viraria uma
 * segunda falha logo na abertura — em cima de um usuário que acabou de perder
 * trabalho.
 */
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

/** Apaga o rascunho. Silencioso: não haver o que apagar é o caso normal. */
export async function discardDraft(): Promise<void> {
  await rm(file(), { force: true }).catch(() => undefined)
}
