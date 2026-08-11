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
 * quando o modelo evoluir na Fase 4.
 */
export const SDOC_FORMAT = 'sdoc'
export const SDOC_VERSION = 1

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

  return { page, doc: parsed.data.doc }
}

export function isDocumentFile(text: string): boolean {
  return text.trimStart().startsWith('{') && text.includes(`"${SDOC_FORMAT}"`)
}
