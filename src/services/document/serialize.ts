import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import {
  documentCommentSchema,
  documentNotesSchema,
  documentPropertiesSchema,
  pageSetupSchema,
  sectionSetupSchema,
  styleSheetSchema,
} from '@shared/schemas.js'
import {
  DEFAULT_PAGE_SETUP,
  isValidMargins,
  type DocumentComment,
  type DocumentModel,
  type DocumentNode,
  type DocumentNotes,
  type DocumentProperties,
  type NoteNumbering,
} from './model.js'
import { LEGACY_STYLES, type StyleSheet } from './styles.js'

/**
 * Formato interno `.sdoc`.
 *
 * É um JSON: o modelo do documento gravado como está. Não substitui o DOCX —
 * serve para salvar e reabrir **sem perda nenhuma**, o que o `.txt` não
 * permite. Imagens vão embutidas como data URI; se isso vier a pesar, o
 * container pode virar ZIP sem que nada fora deste arquivo mude.
 *
 * O campo `version` existe para que um arquivo gravado hoje continue legível
 * quando o modelo evoluir. Cada versão que muda a forma do documento ganha uma
 * migração em `migrate`, aplicada na leitura:
 *
 * - **2** — a imagem deixou de ser bloco e passou a morar dentro do parágrafo,
 *   como no Word. Editá-la como bloco partia o parágrafo em volta.
 * - **3** — o documento passou a carregar os seus **estilos** (`styles.ts`). Um
 *   arquivo da versão 2 não os tem, e recebe `LEGACY_STYLES` na leitura: são a
 *   aparência que o editor já desenhava, medida por medida, para que o documento
 *   antigo abra idêntico.
 * - **4** — o bloco passou a carregar só a formatação **direta**; o herdado vem
 *   dos estilos. Os blocos de um arquivo anterior continuam achatados, e a
 *   leitura os marca (`flattened`) para que a gravação em DOCX os compare com
 *   uma leitura achatada do original. Os nós não são tocados: desachatar exigiria
 *   o `styles.xml` de cada um, e o achatado desenha igual.
 * - **5** — o leitor do `.docx` passou a produzir marcadores, campos, links
 *   internos e sumário. O rascunho anterior não os tem nos nós, e a leitura o
 *   marca (`beforeReferences`) pelo mesmo motivo da versão 4.
 * - **6** — o documento passou a ter **seções**: `sections` leva as anteriores
 *   à última, e o parágrafo que encerra cada uma leva `sectionBreak`. O
 *   rascunho anterior não tem nem uma coisa nem outra — a página dele é a do
 *   documento inteiro —, e a leitura o marca (`beforeSections`) pelo mesmo
 *   motivo da versão 4.
 * - **7** — o documento passou a ter **comentários**: `comments` leva o corpo
 *   de cada um, e o texto leva as pontas da âncora (`commentStart` e
 *   `commentEnd`). O rascunho anterior não tem as pontas, e a leitura o marca
 *   (`beforeComments`) pelo mesmo motivo da versão 4.
 * - **8** — o texto passou a levar as **revisões**: as marcas `insertion` e
 *   `deletion`, a revisão da marca de parágrafo (`markRevision`) e a da linha
 *   de tabela (`rowRevision`); `trackChanges` é o interruptor do documento. O
 *   rascunho anterior não as tem — o inserido era texto comum e o excluído não
 *   aparecia —, e a leitura o marca (`beforeRevisions`) pelo mesmo motivo da
 *   versão 4.
 * - **9** — o texto passou a levar as **notas** de rodapé e de fim: a
 *   referência é o nó `noteRef`, com o corpo da nota dentro, e `notes` leva a
 *   numeração do documento. O rascunho anterior não tem a referência, e a
 *   leitura o marca (`beforeNotes`) pelo mesmo motivo da versão 4.
 * - **10** — as **propriedades** do documento: título, assunto, autor…, em
 *   `properties`. O rascunho anterior simplesmente não as tem, e não ganha
 *   marca: na gravação em DOCX a ausência é "deixe as do arquivo como estão".
 * - **11** — as **equações**: o nó `math`, com o OMML do arquivo dentro e o
 *   MathML que a tela desenha. O rascunho anterior não tem o nó — a equação
 *   ficava escondida no parágrafo —, e a leitura o marca (`beforeMath`) pelo
 *   mesmo motivo da versão 4.
 */
export const SDOC_FORMAT = 'sdoc'
export const SDOC_VERSION = 11

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
  // Só presente quando verdadeiro — ver `DocumentModel.flattened`.
  flattened: z.boolean().optional(),
  // Só presente quando verdadeiro — ver `DocumentModel.beforeReferences`.
  beforeReferences: z.boolean().optional(),
  // Ver `DocumentModel.sections` e `beforeSections`.
  sections: z.array(sectionSetupSchema).max(10_000).optional(),
  beforeSections: z.boolean().optional(),
  outsideBookmarks: z.array(z.string()).optional(),
  // Ver `DocumentModel.comments` e `beforeComments`.
  comments: z.array(documentCommentSchema).max(100_000).optional(),
  beforeComments: z.boolean().optional(),
  // Ver `DocumentModel.trackChanges` e `beforeRevisions`.
  trackChanges: z.boolean().optional(),
  beforeRevisions: z.boolean().optional(),
  // Ver `DocumentModel.notes` e `beforeNotes`.
  notes: documentNotesSchema.optional(),
  beforeNotes: z.boolean().optional(),
  // Ver `DocumentModel.beforeMath`.
  beforeMath: z.boolean().optional(),
  // Ver `DocumentModel.properties`.
  properties: documentPropertiesSchema.optional(),
})

export function serializeDocument(model: DocumentModel): string {
  return JSON.stringify(
    {
      format: SDOC_FORMAT,
      version: SDOC_VERSION,
      page: model.page,
      doc: model.doc,
      // No envelope, como o sidecar os põe ao abrir um `.docx`: os nós, e a
      // impressão digital deles, ficam como estavam.
      styles: model.styles,
      ...(model.flattened === true ? { flattened: true } : {}),
      ...(model.beforeReferences === true ? { beforeReferences: true } : {}),
      ...(model.sections === undefined || model.sections.length === 0 ? {} : { sections: model.sections }),
      ...(model.beforeSections === true ? { beforeSections: true } : {}),
      ...(model.outsideBookmarks === undefined ? {} : { outsideBookmarks: model.outsideBookmarks }),
      ...(model.comments === undefined || model.comments.length === 0 ? {} : { comments: model.comments }),
      ...(model.beforeComments === true ? { beforeComments: true } : {}),
      ...(model.trackChanges === undefined ? {} : { trackChanges: model.trackChanges }),
      ...(model.beforeRevisions === true ? { beforeRevisions: true } : {}),
      ...(model.notes === undefined ? {} : { notes: model.notes }),
      ...(model.beforeNotes === true ? { beforeNotes: true } : {}),
      ...(model.beforeMath === true ? { beforeMath: true } : {}),
      ...(model.properties === undefined ? {} : { properties: model.properties }),
    },
    null,
    2,
  )
}

/** Arquivo corrompido ou de versão futura produz uma frase que a pessoa entenda, e não um erro de JSON. */
export function parseDocument(text: string, language: Language = Language.Portuguese): DocumentModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.corrupt'))
  }

  const parsed = sdocSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.invalid'))
  }

  if (parsed.data.version > SDOC_VERSION) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.newerVersion'))
  }

  // Margens inválidas não impedem a leitura: o documento é recuperado com a
  // configuração padrão, porque o texto do usuário vale mais que o layout.
  const page = isValidMargins(parsed.data.page) ? parsed.data.page : DEFAULT_PAGE_SETUP
  const sections = (parsed.data.sections ?? []).map((section) =>
    isValidMargins(section) ? section : { ...DEFAULT_PAGE_SETUP, id: section.id },
  )

  return {
    page,
    doc: migrate(parsed.data.doc, parsed.data.version),
    styles: migrateStyles(parsed.data.styles, parsed.data.version),
    ...(parsed.data.version < 4 || parsed.data.flattened === true ? { flattened: true } : {}),
    ...(parsed.data.version < 5 || parsed.data.beforeReferences === true ? { beforeReferences: true } : {}),
    ...(sections.length > 0 ? { sections } : {}),
    ...(parsed.data.version < 6 || parsed.data.beforeSections === true ? { beforeSections: true } : {}),
    ...(parsed.data.outsideBookmarks === undefined ? {} : { outsideBookmarks: parsed.data.outsideBookmarks }),
    ...(parsed.data.comments === undefined ? {} : { comments: parsed.data.comments.map(commentOf) }),
    ...(parsed.data.version < 7 || parsed.data.beforeComments === true ? { beforeComments: true } : {}),
    ...(parsed.data.trackChanges === undefined ? {} : { trackChanges: parsed.data.trackChanges }),
    ...(parsed.data.version < 8 || parsed.data.beforeRevisions === true ? { beforeRevisions: true } : {}),
    ...(parsed.data.notes === undefined ? {} : { notes: notesOf(parsed.data.notes) }),
    ...(parsed.data.version < 9 || parsed.data.beforeNotes === true ? { beforeNotes: true } : {}),
    ...(parsed.data.version < 11 || parsed.data.beforeMath === true ? { beforeMath: true } : {}),
    ...(parsed.data.properties === undefined ? {} : { properties: propertiesOf(parsed.data.properties) }),
  }
}

/** As propriedades sem as chaves ausentes — `exactOptionalPropertyTypes`. */
export function propertiesOf(raw: z.infer<typeof documentPropertiesSchema>): DocumentProperties {
  return Object.fromEntries(
    Object.entries(raw).filter(([, value]) => value !== undefined),
  ) as DocumentProperties
}

/** A numeração das notas sem as chaves ausentes — `exactOptionalPropertyTypes`. */
export function notesOf(raw: z.infer<typeof documentNotesSchema>): DocumentNotes {
  const numbering = (pr: NonNullable<typeof raw.footnotePr>): NoteNumbering => ({
    ...(pr.numFmt === undefined ? {} : { numFmt: pr.numFmt }),
    ...(pr.start === undefined ? {} : { start: pr.start }),
    ...(pr.restart === undefined ? {} : { restart: pr.restart }),
    ...(pr.pos === undefined ? {} : { pos: pr.pos }),
  })
  return {
    ...(raw.footnotePr === undefined ? {} : { footnotePr: numbering(raw.footnotePr) }),
    ...(raw.endnotePr === undefined ? {} : { endnotePr: numbering(raw.endnotePr) }),
  }
}

/** O comentário do envelope sem as chaves ausentes — `exactOptionalPropertyTypes`. */
export function commentOf(raw: z.infer<typeof documentCommentSchema>): DocumentComment {
  return {
    id: raw.id,
    author: raw.author,
    date: raw.date,
    paragraphs: raw.paragraphs,
    done: raw.done,
    ...(raw.parentId === undefined ? {} : { parentId: raw.parentId }),
    ...(raw.initials === undefined ? {} : { initials: raw.initials }),
    ...(raw.paraId === undefined ? {} : { paraId: raw.paraId }),
    ...(raw.rich === true ? { rich: true } : {}),
  }
}

/** Traz um documento gravado por uma versão anterior do formato para a atual. */
function migrate(doc: DocumentNode, version: number): DocumentNode {
  return version < 2 ? wrapLooseImages(doc) : doc
}

/**
 * `LEGACY_STYLES`, e não `BUILTIN_STYLES`: o arquivo antigo tem de reabrir com a
 * paginação de antes. Pela versão, e não pela presença do campo: um arquivo da
 * versão 2 com `styles` é um arquivo remendado.
 */
function migrateStyles(styles: StyleSheet | undefined, version: number): StyleSheet {
  return version < 3 || styles === undefined ? LEGACY_STYLES : styles
}

/** Só nos nós de texto a imagem inline tem lugar. */
const TEXTBLOCKS = new Set(['paragraph', 'heading', 'codeBlock'])

/**
 * Embrulha num parágrafo cada imagem que a versão 1 deixou solta entre blocos.
 * O Tiptap monta o conteúdo sem validar, e sem isto ela sobreviveria por acaso
 * até o primeiro caminho que valide.
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
