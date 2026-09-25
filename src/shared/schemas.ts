import { z } from 'zod'
import type { DocumentNode } from '@services/document/model.js'
import { Language, LANGUAGES } from './i18n/language.js'
import { Theme } from './types.js'

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
  /** Pilha de CSS, como o leitor a resolveu. */
  fontFamily: z.string().max(200).optional(),
  /** A peça abre linha nova: no arquivo ela começa outro parágrafo. */
  line: z.boolean().optional(),
  /**
   * Onde a peça mora no arquivo: a relação, o parágrafo e a peça nele.
   *
   * Sem declarar o campo, o zod o **descartava em silêncio** — e o texto
   * digitado no cabeçalho voltaria para a tela e não para o `.docx`.
   */
  pid: z.string().max(120).optional(),
})

/**
 * Cabeçalho ou rodapé vindo de um documento do Word.
 *
 * Três colunas e um filete — o modelo que o Word sempre usou, e que cobre
 * quase todo cabeçalho corporativo.
 *
 * O texto das peças que têm endereço é editável; todo o resto da parte OOXML
 * volta intacto para o arquivo. Ver docs/02-docx-cirurgico.md.
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
  kind: z.enum(['image', 'text', 'rule']),
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
  /** Onde a caixa mora no arquivo, quando o texto dela é editável. */
  bid: z.string().max(120).optional(),
  /** Moldura e preenchimento, quando o leitor soube reproduzi-los. */
  fill: z.string().max(32).optional(),
  line: z.string().max(32).optional(),
  lineWidthPt: z.number().min(0).max(200).optional(),
  dash: z.boolean().optional(),
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

/**
 * As preferências de edição, validadas.
 *
 * Mora aqui, e não no contrato de IPC, porque o mesmo schema serve em três
 * pontos: a leitura do arquivo onde o main as guarda, o pedido do renderer e o
 * aviso que volta para ele. Três definições divergiriam, e a divergência
 * apareceria como um menu marcado que o editor não obedece.
 *
 * Os `default` são o que permite abrir uma instalação antiga: o arquivo gravado
 * antes desta versão não tem chave nenhuma destas.
 */
export const editorPreferencesSchema = z.object({
  spellcheck: z.boolean().default(true),
  invisibleCharacters: z.boolean().default(false),
  typography: z.boolean().default(true),
  // O `default` aqui é só a rede de segurança do parse. Na primeira execução
  // quem escolhe é o sistema operacional, e isso acontece em
  // `src/main/preferences.ts`, que sabe distinguir "chave ausente" de "chave
  // gravada com este valor" — distinção que um `default` apaga.
  language: z.enum(LANGUAGES).default(Language.Portuguese),
  theme: z.enum([Theme.System, Theme.Light, Theme.Dark]).default(Theme.System),
  readingMode: z.boolean().default(false),
  showToolbar: z.boolean().default(true),
  showStatusBar: z.boolean().default(true),
  zoom: z.number().int().min(50).max(200).default(100),
  zoomFit: z.boolean().default(false),
})

/**
 * O remendo: uma ou mais chaves, e **só** as que vieram.
 *
 * Escrito à mão em vez de `editorPreferencesSchema.partial()`, e o teste de
 * contrato existe por causa disto: `.partial()` torna as chaves opcionais mas
 * **mantém os `default`**, então um pedido de "mostrar marcas" voltava do parse
 * com as outras duas chaves preenchidas com o padrão — e desligar a ortografia era
 * desfeito no clique seguinte em qualquer outra chave.
 */
export const editorPreferencesPatchSchema = z.object({
  spellcheck: z.boolean().optional(),
  invisibleCharacters: z.boolean().optional(),
  typography: z.boolean().optional(),
  language: z.enum(LANGUAGES).optional(),
  theme: z.enum([Theme.System, Theme.Light, Theme.Dark]).optional(),
  readingMode: z.boolean().optional(),
  showToolbar: z.boolean().optional(),
  showStatusBar: z.boolean().optional(),
  zoom: z.number().int().min(50).max(200).optional(),
  zoomFit: z.boolean().optional(),
})

/**
 * O que o Chromium conta sobre o ponto onde o botão direito foi clicado.
 *
 * Os tetos não são burocracia: `dictionarySuggestions` alimenta itens de menu, e
 * uma lista longa sairia da tela. O Chromium manda cinco.
 */
export const contextMenuTargetSchema = z.object({
  x: z.number().int().min(0).max(100_000),
  y: z.number().int().min(0).max(100_000),
  editable: z.boolean(),
  misspelledWord: z.string().max(200),
  dictionarySuggestions: z.array(z.string().max(200)).max(10),
  canCut: z.boolean(),
  canCopy: z.boolean(),
  canPaste: z.boolean(),
})

/**
 * Os estilos do documento, validados.
 *
 * Atravessam o IPC em dois sentidos: chegam do sidecar ao abrir um `.docx` e
 * voltam do renderer dentro do `.sdoc` ao salvar. O mesmo schema nos dois pontos
 * porque é o mesmo dado — e porque estilo malformado não pode virar tela: é ele
 * que a entrega seguinte vai usar para desenhar.
 *
 * Nada aqui é `strict`: um `w:pPr` de estilo tem dezenas de propriedades, e o
 * leitor lê as que sabe. Recusar o documento por causa de uma chave nova seria
 * trocar uma tela incompleta por nenhuma tela.
 */
const lineSpacingSchema = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('multiple'), factor: z.number().min(0).max(100) }),
  z.object({ kind: z.literal('exact'), pt: z.number().min(0).max(2000) }),
  z.object({ kind: z.literal('atLeast'), pt: z.number().min(0).max(2000) }),
])

const styleParagraphSchema = z.object({
  textAlign: z.string().max(16).optional(),
  indentMm: z.number().optional(),
  indentRightMm: z.number().optional(),
  firstLineMm: z.number().optional(),
  spaceBefore: z.number().optional(),
  spaceAfter: z.number().optional(),
  lineSpacing: lineSpacingSchema.optional(),
  keepNext: z.boolean().optional(),
  keepLines: z.boolean().optional(),
  widowControl: z.boolean().optional(),
  pageBreakBefore: z.boolean().optional(),
  contextualSpacing: z.boolean().optional(),
  outlineLevel: z.number().int().min(0).max(8).optional(),
  background: z.string().max(32).optional(),
})

const styleCharacterSchema = z.object({
  fontFamily: z.string().max(200).optional(),
  fontSize: z.string().max(16).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strike: z.boolean().optional(),
  allCaps: z.boolean().optional(),
  smallCaps: z.boolean().optional(),
  verticalAlign: z.string().max(16).optional(),
  color: z.string().max(32).optional(),
  highlight: z.string().max(32).optional(),
})

/**
 * Um estilo.
 *
 * Os três interruptores têm padrão porque são derivados da presença de um
 * elemento no arquivo: um `.sdoc` editado à mão sem eles é um estilo que não
 * esconde nem recomenda nada, e não um documento inválido.
 */
const styleDefinitionSchema = z.object({
  id: z.string().min(1).max(120),
  name: z.string().min(1).max(200),
  type: z.enum(['paragraph', 'character']),
  qFormat: z.boolean().default(false),
  hidden: z.boolean().default(false),
  custom: z.boolean().default(false),
  basedOn: z.string().max(120).optional(),
  next: z.string().max(120).optional(),
  link: z.string().max(120).optional(),
  uiPriority: z.number().int().min(0).max(1000).optional(),
  paragraph: styleParagraphSchema.optional(),
  character: styleCharacterSchema.optional(),
})

export const styleSheetSchema = z.object({
  defaults: z.object({
    paragraph: styleParagraphSchema.default({}),
    character: styleCharacterSchema.default({}),
    /** O estilo que vale sem `w:pStyle`; `null` quando o documento não marca nenhum. */
    paragraphStyleId: z.string().max(120).nullable().default(null),
    characterStyleId: z.string().max(120).nullable().default(null),
  }),
  // O teto é a rede contra arquivo patológico, e não um limite de projeto: um
  // documento do Word com estilo para cada variante de tabela passa dos 400.
  styles: z
    .record(z.string().max(120), styleDefinitionSchema)
    .refine((styles) => Object.keys(styles).length <= 4000, 'estilos demais'),
})
