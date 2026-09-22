import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import { pageSetupSchema, styleSheetSchema } from '@shared/schemas.js'
import { DEFAULT_PAGE_SETUP, isValidMargins, type DocumentModel, type DocumentNode } from './model.js'
import { BUILTIN_STYLES, type StyleSheet } from './styles.js'

/**
 * Formato interno `.sdoc`.
 *
 * É um JSON: o modelo do documento gravado como está. Não substitui o DOCX —
 * serve para que a Fase 2 possa salvar e reabrir **sem perda nenhuma**, o que
 * o `.txt` não permite. Imagens vão embutidas como data URI; se isso vier a
 * pesar, o container pode virar ZIP sem que nada fora deste arquivo mude.
 *
 * O campo `version` existe para que um arquivo gravado hoje continue legível
 * quando o modelo evoluir. Cada versão que muda a forma do documento ganha uma
 * migração em `migrate`, aplicada na leitura:
 *
 * - **2** — a imagem deixou de ser bloco e passou a morar dentro do parágrafo,
 *   como no Word. Editá-la como bloco partia o parágrafo em volta.
 * - **3** — o documento passou a carregar os seus **estilos** (`styles.ts`). Um
 *   arquivo da versão 2 não os tem, e recebe `BUILTIN_STYLES` na leitura: são a
 *   aparência que o editor já desenhava, medida por medida, para que o documento
 *   antigo abra idêntico.
 */
export const SDOC_FORMAT = 'sdoc'
export const SDOC_VERSION = 3

/** O conteúdo é validado só na forma; a estrutura fina é do ProseMirror. */
const documentNodeSchema: z.ZodType<DocumentNode> = z.looseObject({
  type: z.string(),
})

const sdocSchema = z.object({
  format: z.literal(SDOC_FORMAT),
  version: z.number().int().positive(),
  page: pageSetupSchema,
  doc: documentNodeSchema,
  // Opcional porque a versão 2 não tem estilos: quem decide o que fazer com a
  // ausência é `migrate`, e não o schema.
  styles: styleSheetSchema.optional(),
})

export function serializeDocument(model: DocumentModel): string {
  return JSON.stringify(
    {
      format: SDOC_FORMAT,
      version: SDOC_VERSION,
      page: model.page,
      doc: model.doc,
      // No envelope, e não dentro do documento: é o mesmo lugar onde o sidecar os
      // põe ao abrir um `.docx`, e é o que mantém os nós — e a impressão digital
      // deles — como estavam.
      styles: model.styles,
    },
    null,
    2,
  )
}

/**
 * Lê um `.sdoc`.
 *
 * Um arquivo corrompido ou de versão futura precisa produzir uma frase que o
 * usuário entenda — não um erro de JSON. Perder o arquivo por não conseguir
 * explicar o problema seria o pior desfecho.
 */
export function parseDocument(text: string, language: Language = Language.Portuguese): DocumentModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      translate(language, 'errors.document.corrupt'),
    )
  }

  const parsed = sdocSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      translate(language, 'errors.document.invalid'),
    )
  }

  if (parsed.data.version > SDOC_VERSION) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      translate(language, 'errors.document.newerVersion'),
    )
  }

  // Margens inválidas não impedem a leitura: o documento é recuperado com a
  // configuração padrão, porque o texto do usuário vale mais que o layout.
  const page = isValidMargins(parsed.data.page) ? parsed.data.page : DEFAULT_PAGE_SETUP

  return {
    page,
    doc: migrate(parsed.data.doc, parsed.data.version),
    styles: migrateStyles(parsed.data.styles, parsed.data.version),
  }
}

/** Traz um documento gravado por uma versão anterior do formato para a atual. */
function migrate(doc: DocumentNode, version: number): DocumentNode {
  return version < 2 ? wrapLooseImages(doc) : doc
}

/**
 * Os estilos de um arquivo que não os tinha.
 *
 * `BUILTIN_STYLES` reproduzem a aparência com que o editor já desenhava o
 * documento — Times New Roman 12 pt, entrelinha 1,5, os títulos como estão hoje —,
 * então o arquivo antigo abre **idêntico**. Dar-lhe outro padrão seria mudar, sem
 * pedir, a paginação de um trabalho já entregue.
 *
 * Pela versão, e não pela presença do campo: um arquivo da versão 2 com um
 * `styles` qualquer não é um arquivo de estilos, é um arquivo remendado.
 *
 * **Atenção para o dia em que o documento novo mudar de padrão** (o dono já
 * decidiu: Calibri 11 pt, entrelinha 1,08, 8 pt depois). Aí `BUILTIN_STYLES`
 * passa a ser o padrão novo, e este lugar **não** pode seguir junto: o arquivo da
 * versão 2 foi escrito por um editor que desenhava Times New Roman 12 pt com
 * entrelinha 1,5, e é essa a aparência que ele tem de reencontrar. A tabela de
 * hoje vira a tabela dos arquivos antigos, com nome próprio, e esta função aponta
 * para ela.
 */
function migrateStyles(styles: StyleSheet | undefined, version: number): StyleSheet {
  return version < 3 || styles === undefined ? BUILTIN_STYLES : styles
}

/**
 * Nós que guardam texto e marcas, e não outros blocos: só neles a imagem, hoje
 * inline, tem lugar.
 */
const TEXTBLOCKS = new Set(['paragraph', 'heading', 'codeBlock'])

/**
 * Embrulha num parágrafo cada imagem que a versão 1 deixou solta num contêiner
 * de blocos — entre os parágrafos, numa célula, num item de lista.
 *
 * Sem isto a imagem abre e aparece, mas num documento que o schema não aceita:
 * o TipTap monta o conteúdo sem validar, e ela sobrevive por acaso até o
 * primeiro caminho que valide. Aqui ela ganha o lugar que o schema exige, e
 * nada do que ela era se perde.
 */
function wrapLooseImages(node: DocumentNode): DocumentNode {
  if (node.content === undefined || TEXTBLOCKS.has(node.type)) return node

  return {
    ...node,
    content: node.content.map((child) =>
      child.type === 'image' ? { type: 'paragraph', content: [child] } : wrapLooseImages(child),
    ),
  }
}

export function isDocumentFile(text: string): boolean {
  return text.trimStart().startsWith('{') && text.includes(`"${SDOC_FORMAT}"`)
}
