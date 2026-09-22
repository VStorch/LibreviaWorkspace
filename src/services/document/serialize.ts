import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { pageSetupSchema } from '@shared/schemas.js'
import { DEFAULT_PAGE_SETUP, isValidMargins, type DocumentModel, type DocumentNode } from './model.js'

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
 */
export const SDOC_FORMAT = 'sdoc'
export const SDOC_VERSION = 2

/** O conteúdo é validado só na forma; a estrutura fina é do ProseMirror. */
const documentNodeSchema: z.ZodType<DocumentNode> = z.looseObject({
  type: z.string(),
})

const sdocSchema = z.object({
  format: z.literal(SDOC_FORMAT),
  version: z.number().int().positive(),
  page: pageSetupSchema,
  doc: documentNodeSchema,
})

export function serializeDocument(model: DocumentModel): string {
  return JSON.stringify(
    { format: SDOC_FORMAT, version: SDOC_VERSION, page: model.page, doc: model.doc },
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
export function parseDocument(text: string): DocumentModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este arquivo não pôde ser lido: o conteúdo está corrompido ou não é um documento válido.',
    )
  }

  const parsed = sdocSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este arquivo não é um documento válido deste aplicativo.',
    )
  }

  if (parsed.data.version > SDOC_VERSION) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este documento foi criado por uma versão mais recente do aplicativo. Atualize para abri-lo.',
    )
  }

  // Margens inválidas não impedem a leitura: o documento é recuperado com a
  // configuração padrão, porque o texto do usuário vale mais que o layout.
  const page = isValidMargins(parsed.data.page) ? parsed.data.page : DEFAULT_PAGE_SETUP

  return { page, doc: migrate(parsed.data.doc, parsed.data.version) }
}

/** Traz um documento gravado por uma versão anterior do formato para a atual. */
function migrate(doc: DocumentNode, version: number): DocumentNode {
  return version < 2 ? wrapLooseImages(doc) : doc
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
