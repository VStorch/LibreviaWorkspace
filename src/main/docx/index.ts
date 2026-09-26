/**
 * DOCX no processo main: conversa com o sidecar e guarda os bytes originais.
 *
 * Os bytes ficam **aqui**, e não no sidecar, por decisão de desenho
 * (docs/02-docx-cirurgico.md): o sidecar é sem estado, então a morte dele não
 * custa a capacidade de gravar cirurgicamente. Guardá-los lá quebraria a
 * promessa da Fase 3.5.
 */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { SDOC_FORMAT, SDOC_VERSION } from '@services/document/serialize.js'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import { styleSheetSchema } from '@shared/schemas.js'
import type { LossInventory } from '@shared/types.js'
import { normalizePath } from '../fs/paths.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/**
 * Os rótulos do inventário, já cortados nos limites do contrato de IPC.
 *
 * `ipc.ts` recusa mais de 50 rótulos por categoria e mais de 300 caracteres em
 * cada um, e o registro passou a executar esse schema também na resposta. Sem o
 * corte aqui, um documento patológico — dezenas de medidas inválidas distintas —
 * deixaria de abrir por causa do **aviso**, e não do conteúdo. Cortar a lista de
 * avisos é o desfecho certo; recusar o arquivo por causa dela, não.
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
 * O resultado de abrir, conferido antes de virar documento na tela.
 *
 * `page` e `doc` atravessam como desconhecidos — quem os valida é o schema do
 * `.sdoc`, na ponta do renderer, e o do ProseMirror depois dele. Os **estilos**
 * não: eles são conferidos aqui, no primeiro ponto em que entram no programa,
 * porque é este o único lugar por onde passam antes de serem gravados no arquivo
 * do usuário. Um estilo malformado que chegasse ao `.sdoc` ficaria lá.
 */
const openResultSchema = z.object({
  model: z.object({ page: z.unknown(), doc: z.unknown(), styles: styleSheetSchema }),
  inventory: inventorySchema,
})

const saveResultSchema = z.object({
  inventory: inventorySchema,
  preservedBlocks: z.number().int().nonnegative(),
  rewrittenBlocks: z.number().int().nonnegative(),
})

/**
 * Os bytes do documento aberto, como estavam no disco na hora de abrir.
 *
 * Uma entrada só: o aplicativo edita um documento por vez. E é de propósito que
 * a fonte de verdade seja *o arquivo como você o abriu*, e não como ele está no
 * disco agora — que pode ter mudado por outra mão no meio da edição.
 */
let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedDocx(): void {
  openedOriginal = null
}

/**
 * Reata o vínculo com o pacote original depois de uma recuperação.
 *
 * Depois de uma queda, o modelo volta do rascunho mas os bytes originais não —
 * eles moravam na memória do processo que morreu. Sem relê-los, salvar por cima
 * do `.docx` recuperado seria recusado, e o usuário ficaria com o trabalho na
 * tela sem poder gravá-lo onde ele estava.
 */
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
  /** O modelo já no envelope `.sdoc`, para o renderer seguir por um caminho só. */
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

  // O caminho entra normalizado: é assim que ele volta do `authorizePath` e é
  // assim que o `origin` chega na gravação. Guardado cru, um `/tmp/./a.docx`
  // não se reconheceria no `/tmp/a.docx` da gravação, e o `.docx` aberto seria
  // gravado por cima do pacote mínimo — todo `oid` descartado com ele.
  openedOriginal = { path: normalizePath(path), bytes }

  return {
    // As constantes, e não o literal: o envelope sai daqui rotulado com a versão
    // do formato, e um `1` fixo faria todo documento vindo do Word passar por
    // migrações que ele não precisa quando o formato mudasse.
    content: JSON.stringify({ format: SDOC_FORMAT, version: SDOC_VERSION, ...parsed.data.model }),
    inventory: parsed.data.inventory,
  }
}

export interface SavedDocx {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
  /** Os bytes que passam a ser o original quando a gravação chegar ao disco. */
  readonly original: Buffer
}

/** De onde o documento veio e para onde vai. */
export interface DocxTarget {
  /**
   * O caminho de que o documento em edição foi carregado — `null` no documento
   * novo. Só serve para comparar com o original guardado: nenhum byte é lido
   * dele.
   */
  readonly origin: string | null
  readonly destination: string
}

/**
 * Grava o modelo por cima do pacote original — ou, no documento que nasceu no
 * editor, por cima de um pacote mínimo criado pelo sidecar.
 *
 * `sdocContent` é o mesmo JSON que o renderer manda para qualquer destino — o
 * envelope é descascado aqui para que o renderer não precise saber que DOCX
 * existe.
 *
 * O original só vale para o documento que saiu dele. Antes, qualquer gravação
 * em DOCX usava os bytes guardados, e um documento novo criado depois de abrir
 * um `.docx` sairia com os cabeçalhos, os estilos e as notas do outro arquivo.
 * Agora a origem do documento em edição precisa ser o caminho do original.
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
    // `flatten` escolhe a leitura de referência do sidecar: o rascunho antigo
    // traz blocos achatados, e só uma leitura achatada os reconhece. Os estilos
    // vão sempre, para que o modificado e o criado cheguem a `word/styles.xml`
    // (`StyleWriter.cs`) — também do rascunho antigo, que deixa criar e aplicar
    // estilo como qualquer outro documento: gravá-los é seguro ali, porque os
    // blocos achatados carregam os valores como formatação direta.
    {
      page: model.page,
      doc: model.doc,
      ...(model.flattened ? { flatten: true } : {}),
      ...(model.beforeReferences ? { beforeReferences: true } : {}),
      ...(model.styles === undefined ? {} : { styles: model.styles }),
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
    // `includes` porque o sidecar já declara a mesma perda quando a faixa tinha
    // texto para gravar e a relação não existia no pacote mínimo: a frase é uma
    // só, e repetida seriam dois avisos na tela para um problema.
    if (hasForeignBands(model.page) && !lost.includes(t('errors.docx.foreignBands')))
      lost.push(t('errors.docx.foreignBands'))
    // Rede de proteção: um modelo com `oid` foi numerado contra um pacote que
    // não está aqui. Isso é defeito — o vínculo com o original se perdeu no
    // caminho —, e o que sai é o pacote mínimo, sem os estilos, as notas nem os
    // comentários do arquivo de origem. Dito em voz alta, porque perda calada é
    // o pior defeito que este programa pode ter.
    if (hasOid(model.doc)) lost.push(t('errors.docx.originPackage'))
  }

  return {
    bytes: reply.binary,
    inventory: { ...inventory, lost },
    // O original que segue adiante é o pacote de partida desta gravação, e não
    // os bytes gravados. No documento novo isso é o **pacote mínimo**: com os
    // bytes gravados no lugar dele, cada gravação partia do resultado da
    // anterior e somava o que já estava lá — uma definição de numeração e uma
    // cópia de cada imagem por gravação. Partindo sempre do mesmo pacote, e sem
    // nenhum bloco com `oid`, nada se perde e o resultado é o mesmo.
    original,
  }
}

/**
 * O original acompanha o documento para onde ele foi gravado.
 *
 * Chamado **depois** que a gravação chegou ao disco: movido antes, uma gravação
 * que falhasse deixaria o original apontando um caminho que o documento não
 * tem, e a gravação seguinte trocaria o pacote do arquivo pelo modelo mínimo.
 *
 * Vale também para destino que não é `.docx`: o `.docx` salvo como `.sdoc` e
 * depois como `.docx` de novo continua sendo o mesmo documento, com os mesmos
 * `oid`, e o pacote de origem continua sendo o dele.
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

/**
 * Algum bloco do modelo veio de um pacote `.docx`?
 *
 * O `oid` é a impressão digital do bloco no arquivo de origem, e o leitor só o
 * põe em bloco que saiu de um. Basta um para que o modelo esteja falando de um
 * pacote que a gravação precisa ter em mãos.
 */
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

/**
 * O pacote mínimo de um documento que nasceu no editor.
 *
 * Criado pelo sidecar, e não guardado aqui como binário: o pacote é código C#
 * validado pelo mesmo SDK que grava (`DocxTemplate`), e a página e os estilos
 * vêm do documento: o `.sdoc` antigo leva a Times com que foi escrito, e o
 * documento novo, a Calibri.
 */
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
} {
  let parsed: unknown
  try {
    parsed = JSON.parse(content)
  } catch {
    throw new AppError(ErrorCode.Internal, t('errors.docx.inconsistentState'))
  }

  // Os estilos são conferidos aqui e seguem para o pacote do documento novo
  // (`docx.create`) e para a gravação: o sidecar os compara com os do original e
  // só toca `word/styles.xml` quando algum mudou — senão ele volta byte a byte.
  const envelope = z
    .object({
      page: z.unknown(),
      doc: z.unknown(),
      styles: styleSheetSchema.optional(),
      flattened: z.boolean().optional(),
      beforeReferences: z.boolean().optional(),
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
  }
}
