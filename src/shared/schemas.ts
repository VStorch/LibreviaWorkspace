import { z } from 'zod'
import type { DocumentNode } from '@services/document/model.js'

/**
 * Schemas usados em mais de um lugar.
 *
 * A configuração de página é validada tanto ao ler um `.sdoc` do disco quanto
 * ao receber um pedido de impressão do renderer. Duas definições divergiriam,
 * e a divergência apareceria como margem errada no papel.
 */

/** Um pedaço de cabeçalho: texto, imagem ou número de página. */
export const bandPieceSchema = z.object({
  kind: z.enum(['text', 'image', 'pageNumber', 'totalPages']),
  text: z.string().max(1000).optional(),
  /** Data URI. Imagem de cabeçalho é pequena — um logotipo, não uma foto. */
  src: z.string().max(4_000_000).optional(),
  width: z.number().int().positive().max(4000).optional(),
  height: z.number().int().positive().max(4000).optional(),
  bold: z.boolean().default(false),
  italic: z.boolean().default(false),
  color: z.string().max(32).optional(),
  fontSize: z.string().max(16).optional(),
})

/**
 * Cabeçalho ou rodapé vindo de um documento do Word.
 *
 * Três colunas e um filete — o modelo que o Word sempre usou, e que cobre
 * quase todo cabeçalho corporativo. É **somente leitura**: o que vai para o
 * arquivo é a parte OOXML original, preservada intacta pela gravação
 * cirúrgica. Ver docs/02-docx-cirurgico.md.
 */
/**
 * Um nó do documento, do jeito que o editor o entende.
 *
 * Conferido só até "é um nó", de propósito. Quem de fato valida é o serializador
 * do ProseMirror, que constrói a partir do schema do editor e **só emite o que
 * ele conhece** — é essa a barreira que impede um documento de mandar marcação
 * para dentro da página. Repetir a lista de tipos de nó aqui a duplicaria num
 * lugar onde ela não pode ser conferida contra o editor, e a cópia mais velha
 * das duas passaria a mandar.
 */
const documentNodeSchema = z.custom<DocumentNode>(
  (value) => typeof value === 'object' && value !== null && typeof (value as DocumentNode).type === 'string',
)

/**
 * Objeto ancorado dentro da faixa.
 *
 * Aberto de propósito: a geometria vem do arquivo e quem a interpreta é
 * `services/document/floating.ts`. Validar campo a campo aqui obrigaria a
 * duplicar essa interpretação no esquema.
 */
const bandFloatSchema = z.object({
  kind: z.enum(['image', 'text']),
  src: z.string().optional(),
  /**
   * O texto de uma caixa, em nós do documento.
   *
   * Aberto como o resto deste esquema, e pela mesma razão: quem o interpreta é
   * o serializador do editor, que só emite o que o schema dele conhece. Sem
   * declarar o campo, o zod o **descartava em silêncio** — e a caixa do título
   * do cabeçalho aparecia na folha com o tamanho certo e vazia por dentro.
   */
  content: z.array(documentNodeSchema).max(200).optional(),
  widthMm: z.number(),
  heightMm: z.number(),
  rotation: z.number(),
  hFrom: z.string(),
  hOffsetMm: z.number().optional(),
  hAlign: z.string().optional(),
  vFrom: z.string(),
  vOffsetMm: z.number().optional(),
  vAlign: z.string().optional(),
  behind: z.boolean(),
  wrap: z.string(),
  dxMm: z.number().optional(),
  dyMm: z.number().optional(),
})

/**
 * Uma célula da grade do cabeçalho.
 *
 * `borders` são as iniciais dos lados que têm risco — `t`, `l`, `b`, `r` — já
 * resolvidos pelo leitor: no OOXML cada lado vem por três caminhos, e refazer
 * essa conta em dois desenhistas é como tela e papel divergem.
 */
const bandCellSchema = z.object({
  pieces: z.array(bandPieceSchema).max(40).default([]),
  width: z.number().min(0).max(1).default(0),
  span: z.number().int().min(1).max(32).default(1),
  rowSpan: z.number().int().min(1).max(32).default(1),
  align: z.string().max(16).optional(),
  borders: z.string().max(4).default(''),
})

export const bandSchema = z.object({
  left: z.array(bandPieceSchema).max(20).default([]),
  center: z.array(bandPieceSchema).max(20).default([]),
  right: z.array(bandPieceSchema).max(20).default([]),
  rule: z.boolean().default(false),
  floats: z.array(bandFloatSchema).max(20).default([]),
  rows: z
    .array(z.object({ cells: z.array(bandCellSchema).max(32).default([]) }))
    .max(32)
    .default([]),
})

export const pageSetupSchema = z.object({
  size: z.enum(['A4', 'Letter']),
  orientation: z.enum(['portrait', 'landscape']),
  margins: z.object({
    top: z.number(),
    right: z.number(),
    bottom: z.number(),
    left: z.number(),
  }),
  // Acrescentados na Fase 3. Opcionais para que documentos gravados antes
  // continuem abrindo — acréscimo compatível não exige nova versão de formato.
  header: z.string().max(500).default(''),
  footer: z.string().max(500).default(''),
  // Acrescentados na Fase 4, pelo mesmo motivo. Quando existem, mandam na
  // exibição: são o cabeçalho real do documento, e o texto acima é o que o
  // usuário digitou num documento criado aqui.
  headerBand: bandSchema.nullable().default(null),
  footerBand: bandSchema.nullable().default(null),
  // Primeira página e páginas pares, quando o documento pede. Opcionais pelo
  // mesmo motivo dos anteriores: `.sdoc` gravado antes daqui não os tem.
  firstHeaderBand: bandSchema.nullable().default(null),
  firstFooterBand: bandSchema.nullable().default(null),
  evenHeaderBand: bandSchema.nullable().default(null),
  evenFooterBand: bandSchema.nullable().default(null),
  /**
   * Distância da faixa à borda do papel, em milímetros.
   *
   * Origem vertical das âncoras de dentro do cabeçalho: elas se dizem relativas
   * ao parágrafo, e o parágrafo do cabeçalho começa aqui.
   */
  headerDistanceMm: z.number().default(12.5),
  footerDistanceMm: z.number().default(12.5),
})
