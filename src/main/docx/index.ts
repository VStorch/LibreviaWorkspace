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
import { styleSheetSchema } from '@shared/schemas.js'
import type { LossInventory } from '@shared/types.js'
import { normalizePath } from '../fs/paths.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'

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
    throw fromFileSystemError(cause, 'leitura')
  }

  const reply = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível ler este documento do Word. O arquivo pode estar danificado.',
      'docx.open fora do contrato',
    )
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
  const original = kept ?? Buffer.from(await createDocx(client, model.page))

  const reply = await client.request(SidecarMethod.DocxSave, model, new Uint8Array(original))
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível gravar o documento do Word. O arquivo original não foi alterado.',
      'docx.save fora do contrato',
    )
  }

  console.info(`[docx] preservados ${parsed.data.preservedBlocks}, reescritos ${parsed.data.rewrittenBlocks}`)

  const inventory = parsed.data.inventory
  const lost = [...inventory.lost]
  if (kept === null) {
    // `includes` porque o sidecar já declara a mesma perda quando a faixa tinha
    // texto para gravar e a relação não existia no pacote mínimo: a frase é uma
    // só, e repetida seriam dois avisos na tela para um problema.
    if (hasForeignBands(model.page) && !lost.includes(FOREIGN_BANDS)) lost.push(FOREIGN_BANDS)
    // Rede de proteção: um modelo com `oid` foi numerado contra um pacote que
    // não está aqui. Isso é defeito — o vínculo com o original se perdeu no
    // caminho —, e o que sai é o pacote mínimo, sem os estilos, as notas nem os
    // comentários do arquivo de origem. Dito em voz alta, porque perda calada é
    // o pior defeito que este programa pode ter.
    if (hasOid(model.doc)) lost.push(ORIGIN_PACKAGE)
  }

  return {
    bytes: reply.binary,
    inventory: lost.length === inventory.lost.length ? inventory : { ...inventory, lost },
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
 * Faixas de cabeçalho e rodapé que vieram de outro `.docx`.
 *
 * Elas só se gravam editando as partes do pacote de onde saíram. Sem esse
 * pacote — o `.sdoc` que um dia foi `.docx`, reaberto do disco — não há onde
 * escrevê-las, e o modelo mínimo sai sem elas. Perda sim, mas dita.
 */
const FOREIGN_BANDS = 'cabeçalho e rodapé do arquivo .docx de origem'

/**
 * O pacote inteiro do arquivo de origem, quando ele não está aqui.
 *
 * Os `oid` do modelo só existem porque um `.docx` foi aberto: cada um aponta um
 * bloco daquele pacote. Sem os bytes dele, a gravação parte do pacote mínimo e
 * nada mais do arquivo de origem — estilos, numeração, notas, comentários,
 * faixas — chega ao destino.
 */
const ORIGIN_PACKAGE = 'estilos, notas, comentários e demais partes do arquivo .docx de origem'

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
 * validado pelo mesmo SDK que grava (`DocxTemplate`), e a página vem da
 * configuração do documento.
 */
async function createDocx(client: SidecarClient, page: unknown): Promise<Uint8Array> {
  const reply = await client.request(SidecarMethod.DocxCreate, page)
  if (reply.binary.length === 0) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível criar o documento do Word. Nada foi gravado.',
      'docx.create sem pacote',
    )
  }
  return reply.binary
}

function unwrapSdoc(content: string): { page: unknown; doc: unknown } {
  let parsed: unknown
  try {
    parsed = JSON.parse(content)
  } catch {
    throw new AppError(ErrorCode.Internal, 'O documento em edição está em estado inconsistente.')
  }

  // Os estilos são conferidos e **não** seguem para o sidecar: nesta entrega o
  // escritor não os usa — `word/styles.xml` volta ao arquivo byte a byte pela
  // gravação cirúrgica, e no documento novo o pacote mínimo já os grava a partir
  // da mesma tabela (`BuiltinStyles.cs`). Conferi-los aqui é o que garante que o
  // modelo em edição continua sendo um modelo válido; mandá-los adiante seria
  // prometer uma gravação que ainda não existe.
  const envelope = z
    .object({ page: z.unknown(), doc: z.unknown(), styles: styleSheetSchema.optional() })
    .safeParse(parsed)
  if (!envelope.success) {
    throw new AppError(ErrorCode.Internal, 'O documento em edição está em estado inconsistente.')
  }

  return { page: envelope.data.page, doc: envelope.data.doc }
}
