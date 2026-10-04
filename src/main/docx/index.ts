/** Os bytes originais ficam **aqui**, e não no sidecar, que é sem estado: a morte dele não custa a gravação cirúrgica. */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { SDOC_FORMAT, SDOC_VERSION } from '@services/document/serialize.js'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import {
  documentCommentSchema,
  documentNotesSchema,
  documentPropertiesSchema,
  styleSheetSchema,
} from '@shared/schemas.js'
import type { LossInventory } from '@shared/types.js'
import { normalizePath } from '../fs/paths.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'
import { MAX_NAME_LENGTH } from '@shared/limits.js'

/**
 * `ipc.ts` recusa mais de 50 rótulos por categoria e mais de 300 caracteres em
 * cada um: sem o corte, um documento patológico deixaria de abrir por causa do
 * aviso.
 */
const inventoryLabels = z
  .array(z.string())
  .default([])
  .transform((labels) => labels.slice(0, 50).map((label) => label.slice(0, 300)))

const inventorySchema = z.object({
  invisible: inventoryLabels,
  lost: inventoryLabels,
  structural: inventoryLabels,
})

/**
 * `page` e `doc` são validados no renderer; os estilos, aqui, o único ponto por
 * onde passam antes de irem para o `.sdoc`.
 */
const openResultSchema = z.object({
  model: z.object({
    page: z.unknown(),
    sections: z.unknown().optional(),
    doc: z.unknown(),
    styles: styleSheetSchema,
    outsideBookmarks: z.array(z.string().max(MAX_NAME_LENGTH)).max(10_000).optional(),
    comments: z.array(documentCommentSchema).max(100_000).optional(),
    trackChanges: z.boolean().optional(),
    notes: documentNotesSchema.optional(),
    properties: documentPropertiesSchema.optional(),
  }),
  inventory: inventorySchema,
})

const saveResultSchema = z.object({
  inventory: inventorySchema,
  preservedBlocks: z.number().int().nonnegative(),
  rewrittenBlocks: z.number().int().nonnegative(),
})

/** Como estavam no disco ao abrir, e não como estão agora, que pode ter mudado por outra mão. */
let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedDocx(): void {
  openedOriginal = null
}

/** Depois de uma queda os bytes originais se perderam com o processo: sem relê-los, salvar seria recusado. */
export async function adoptDocxOriginal(path: string): Promise<boolean> {
  try {
    openedOriginal = { path: normalizePath(path), bytes: await readFile(path) }
    return true
  } catch {
    openedOriginal = null
    return false
  }
}

export interface OpenedDocx {
  readonly content: string
  readonly inventory: LossInventory
}

export async function openDocx(client: SidecarClient, path: string): Promise<OpenedDocx> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  const reply = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.docx.cannotRead'), t('errors.docx.openContract'))
  }

  // Normalizado, como volta do `authorizePath` e chega no `origin`: cru, o `.docx`
  // aberto seria gravado por cima do pacote mínimo.
  openedOriginal = { path: normalizePath(path), bytes }

  return {
    // As constantes, e não o literal: um `1` fixo faria todo documento do Word passar por migrações.
    content: JSON.stringify({ format: SDOC_FORMAT, version: SDOC_VERSION, ...parsed.data.model }),
    inventory: parsed.data.inventory,
  }
}

export interface SavedDocx {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
  readonly original: Buffer
}

export interface DocxTarget {
  /** `null` no documento novo. Só para comparar com o original guardado. */
  readonly origin: string | null
  readonly destination: string
  /** O sidecar grava o rótulo de modelo na parte principal; sem isto, o de documento. */
  readonly template?: boolean
}

/**
 * Por cima do pacote original, ou do pacote mínimo no documento que nasceu no
 * editor. O original só vale se a origem do documento em edição for o caminho
 * dele: senão um documento novo levaria cabeçalhos, estilos e notas de outro.
 */
export async function saveDocx(
  client: SidecarClient,
  sdocContent: string,
  target: DocxTarget,
): Promise<SavedDocx> {
  const model = unwrapSdoc(sdocContent)

  const kept =
    openedOriginal !== null && target.origin !== null && openedOriginal.path === normalizePath(target.origin)
      ? openedOriginal.bytes
      : null
  const original = kept ?? Buffer.from(await createDocx(client, model.page, model.styles))

  const reply = await client.request(
    SidecarMethod.DocxSave,
    // `flatten` escolhe a leitura de referência do sidecar para o rascunho
    // achatado. Os estilos vão sempre, para chegarem a `word/styles.xml`.
    {
      page: model.page,
      doc: model.doc,
      ...(model.flattened ? { flatten: true } : {}),
      ...(model.beforeReferences ? { beforeReferences: true } : {}),
      ...(model.sections === undefined ? {} : { sections: model.sections }),
      ...(model.beforeSections ? { beforeSections: true } : {}),
      // A lista vazia é "todos excluídos"; no rascunho de antes deles, só a marca.
      ...(model.beforeComments ? { beforeComments: true } : { comments: model.comments }),
      ...(model.trackChanges === undefined ? {} : { trackChanges: model.trackChanges }),
      ...(model.beforeRevisions ? { beforeRevisions: true } : {}),
      ...(model.beforeNotes ? { beforeNotes: true } : {}),
      ...(model.beforeMath ? { beforeMath: true } : {}),
      // O sidecar só grava a numeração das notas e as propriedades quando diferem
      // das do arquivo de destino.
      ...(model.notes === undefined ? {} : { notes: model.notes }),
      ...(model.properties === undefined ? {} : { properties: model.properties }),
      ...(model.styles === undefined ? {} : { styles: model.styles }),
      ...(target.template === true ? { template: true } : {}),
    },
    new Uint8Array(original),
  )
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.docx.cannotSave'), t('errors.docx.saveContract'))
  }

  console.info(`[docx] preservados ${parsed.data.preservedBlocks}, reescritos ${parsed.data.rewrittenBlocks}`)

  const inventory = parsed.data.inventory
  const foreignBandsPt = translate(Language.Portuguese, 'errors.docx.foreignBands')
  const lost = inventory.lost.map((item) => (item === foreignBandsPt ? t('errors.docx.foreignBands') : item))
  if (kept === null) {
    // O sidecar já declara a mesma perda quando a relação não existia no pacote
    // mínimo: a frase é uma só.
    if (
      [model.page, ...(Array.isArray(model.sections) ? model.sections : [])].some(hasForeignBands) &&
      !lost.includes(t('errors.docx.foreignBands'))
    )
      lost.push(t('errors.docx.foreignBands'))
    // Rede de proteção: um modelo com `oid` foi numerado contra um pacote que
    // não está aqui, e sai o pacote mínimo. Dito em voz alta, porque perda calada
    // é o pior defeito deste programa.
    if (hasOid(model.doc)) lost.push(t('errors.docx.originPackage'))
  }

  return {
    bytes: reply.binary,
    inventory: { ...inventory, lost },
    // O pacote de partida desta gravação, e não os bytes gravados: partindo do
    // resultado anterior, o documento novo somaria uma cópia de cada imagem por
    // gravação.
    original,
  }
}

/**
 * Chamado **depois** que a gravação chegou ao disco, para uma falha não deixar o
 * original apontando um caminho que o documento não tem. Vale para qualquer
 * destino: o `.docx` salvo como `.sdoc` continua com os mesmos `oid`.
 */
export function followDocxOriginal(
  origin: string | null,
  destination: string,
  saved: SavedDocx | null,
): void {
  const target = normalizePath(destination)
  if (saved !== null) {
    openedOriginal = { path: target, bytes: saved.original }
    return
  }
  if (openedOriginal !== null && origin !== null && openedOriginal.path === normalizePath(origin)) {
    openedOriginal = { path: target, bytes: openedOriginal.bytes }
  }
}

/** O `oid` é a impressão digital do bloco no pacote de origem: basta um. */
function hasOid(doc: unknown): boolean {
  if (typeof doc !== 'object' || doc === null) return false
  const content = (doc as { content?: unknown }).content
  if (!Array.isArray(content)) return false

  return content.some((block) => {
    if (typeof block !== 'object' || block === null) return false
    const attrs = (block as { attrs?: unknown }).attrs
    if (typeof attrs !== 'object' || attrs === null) return false
    return typeof (attrs as { oid?: unknown }).oid === 'string'
  })
}

const bandKeys = [
  'headerBand',
  'footerBand',
  'firstHeaderBand',
  'firstFooterBand',
  'evenHeaderBand',
  'evenFooterBand',
] as const

function hasForeignBands(page: unknown): boolean {
  if (typeof page !== 'object' || page === null) return false
  const record = page as Record<string, unknown>
  return bandKeys.some((key) => record[key] !== null && record[key] !== undefined)
}

/** Criado pelo sidecar (`DocxTemplate`), com a página e os estilos do documento. */
async function createDocx(client: SidecarClient, page: unknown, styles: unknown): Promise<Uint8Array> {
  const reply = await client.request(SidecarMethod.DocxCreate, { page, styles })
  if (reply.binary.length === 0) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.docx.cannotCreate'),
      t('errors.docx.createContract'),
    )
  }
  return reply.binary
}

function unwrapSdoc(content: string): {
  page: unknown
  doc: unknown
  styles: unknown
  flattened: boolean
  beforeReferences: boolean
  sections: unknown
  beforeSections: boolean
  beforeComments: boolean
  comments: unknown[]
  trackChanges: boolean | undefined
  beforeRevisions: boolean
  beforeNotes: boolean
  beforeMath: boolean
  notes: unknown
  properties: unknown
} {
  let parsed: unknown
  try {
    parsed = JSON.parse(content)
  } catch {
    throw new AppError(ErrorCode.Internal, t('errors.docx.inconsistentState'))
  }

  // O sidecar só toca `word/styles.xml` quando algum estilo mudou.
  const envelope = z
    .object({
      page: z.unknown(),
      doc: z.unknown(),
      styles: styleSheetSchema.optional(),
      flattened: z.boolean().optional(),
      beforeReferences: z.boolean().optional(),
      sections: z.unknown().optional(),
      beforeSections: z.boolean().optional(),
      beforeComments: z.boolean().optional(),
      comments: z.array(documentCommentSchema).max(100_000).optional(),
      trackChanges: z.boolean().optional(),
      beforeRevisions: z.boolean().optional(),
      beforeNotes: z.boolean().optional(),
      beforeMath: z.boolean().optional(),
      notes: z.unknown().optional(),
      // Conferidas aqui também: vão para o arquivo do usuário.
      properties: documentPropertiesSchema.optional(),
    })
    .safeParse(parsed)
  if (!envelope.success) {
    throw new AppError(ErrorCode.Internal, t('errors.docx.inconsistentState'))
  }

  return {
    page: envelope.data.page,
    doc: envelope.data.doc,
    styles: envelope.data.styles,
    flattened: envelope.data.flattened === true,
    beforeReferences: envelope.data.beforeReferences === true,
    sections: envelope.data.sections,
    beforeSections: envelope.data.beforeSections === true,
    beforeComments: envelope.data.beforeComments === true,
    comments: envelope.data.comments ?? [],
    trackChanges: envelope.data.trackChanges,
    beforeRevisions: envelope.data.beforeRevisions === true,
    beforeNotes: envelope.data.beforeNotes === true,
    beforeMath: envelope.data.beforeMath === true,
    notes: envelope.data.notes,
    properties: envelope.data.properties,
  }
}
