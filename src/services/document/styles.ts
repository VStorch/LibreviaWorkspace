/**
 * Os estilos do documento, como **dado**.
 *
 * Um `.docx` de verdade não guarda a formatação no parágrafo: guarda o nome de um
 * estilo, e o estilo mora em `word/styles.xml`. Até aqui o leitor resolvia essa
 * cascata no sidecar e achatava o resultado em cada bloco — a tela mostrava o
 * documento certo, mas o documento não tinha estilos: renomear um, ou mudar o
 * "Corpo de texto" de uma vez, não era possível nem em teoria.
 *
 * Esta é a primeira metade do caminho: as definições atravessam o sidecar e
 * chegam ao modelo, **fora dos nós**. Fora dos nós não é detalhe — é o que faz
 * esta entrega ser segura: a impressão digital de um bloco é feita do que está
 * dentro dele, então nada aqui pode fazer a gravação cirúrgica reescrever um
 * bloco que ninguém tocou (ver `src/main/sidecar/fingerprint.test.ts`).
 *
 * As unidades são as mesmas dos atributos do bloco — pontos, milímetros,
 * hexadecimal, `12pt` —, com uma exceção declarada: a **entrelinha**. O bloco
 * guarda o número que o CSS entende, que é o fator do arquivo já multiplicado
 * pela altura natural da fonte; aqui o fator vem cru, porque com herança a fonte
 * pode vir do próprio estilo e a multiplicação só pode acontecer depois de a
 * cascata ser resolvida. Quem multiplica é `line-metrics.ts`.
 *
 * A tela nasce daqui: `style-css.ts` resolve a cascata (`style-cascade.ts`) e
 * gera o CSS de cada estilo, e é esse texto que vai para o editor e para o PDF.
 * Mudar um número aqui muda a aparência — dos documentos que carregam a tabela.
 */

import { Language, translate } from '@shared/i18n/index.js'

/** Como a entrelinha é medida, do jeito que o arquivo a declara. */
export type LineSpacing =
  /** Vezes a altura natural da linha (`w:lineRule="auto"`). */
  | { readonly kind: 'multiple'; readonly factor: number }
  /** A altura é esta, e o que não couber é cortado (`w:lineRule="exact"`). */
  | { readonly kind: 'exact'; readonly pt: number }
  /** A altura é esta ou mais (`w:lineRule="atLeast"`). */
  | { readonly kind: 'atLeast'; readonly pt: number }

/**
 * O que um estilo diz sobre o parágrafo.
 *
 * Todo campo é opcional, e a ausência tem significado: "este estilo não fala
 * disso", que é o silêncio pelo qual o valor herdado passa. Zero é outra coisa —
 * é o estilo dizendo "nenhum espaço aqui".
 *
 * `| undefined` explícito por causa de `exactOptionalPropertyTypes`: o objeto vem
 * do zod, que emite a chave com `undefined`, e sem isso ele não seria atribuível.
 */
export interface StyleParagraphFormat {
  readonly textAlign?: string | undefined
  /** Recuo esquerdo em milímetros, como a régua da configuração de página. */
  readonly indentMm?: number | undefined
  readonly indentRightMm?: number | undefined
  /** Primeira linha: positiva entra, negativa sai (o deslocamento do Word). */
  readonly firstLineMm?: number | undefined
  /** Espaço antes e depois, em pontos — como o Word os mostra. */
  readonly spaceBefore?: number | undefined
  readonly spaceAfter?: number | undefined
  readonly lineSpacing?: LineSpacing | undefined
  readonly keepNext?: boolean | undefined
  readonly keepLines?: boolean | undefined
  readonly pageBreakBefore?: boolean | undefined
  readonly contextualSpacing?: boolean | undefined
  /** Nível na estrutura do documento, de 0 a 8 — o que faz de um estilo título. */
  readonly outlineLevel?: number | undefined
  readonly background?: string | undefined
}

/** O que um estilo diz sobre o texto. */
export interface StyleCharacterFormat {
  /** Pilha de CSS, como o leitor a monta a partir da tabela de fontes. */
  readonly fontFamily?: string | undefined
  /** Medida com unidade, como o atributo do bloco: `12pt`. */
  readonly fontSize?: string | undefined
  readonly bold?: boolean | undefined
  readonly italic?: boolean | undefined
  readonly underline?: boolean | undefined
  readonly strike?: boolean | undefined
  readonly allCaps?: boolean | undefined
  readonly smallCaps?: boolean | undefined
  readonly verticalAlign?: string | undefined
  readonly color?: string | undefined
  readonly highlight?: string | undefined
}

export const StyleType = {
  Paragraph: 'paragraph',
  Character: 'character',
} as const
export type StyleType = (typeof StyleType)[keyof typeof StyleType]

/** Um estilo do documento, como o arquivo o declara — sem herança resolvida. */
export interface StyleDefinition {
  /** O `w:styleId`: o que o parágrafo aponta. É traduzido (`Ttulo1` em português). */
  readonly id: string
  /** O `w:name`: o nome interno, que **não** se traduz (`heading 1` em qualquer idioma). */
  readonly name: string
  readonly type: StyleType
  /** Estilo recomendado, que o Word mostra na galeria (`w:qFormat`). */
  readonly qFormat: boolean
  /** Escondido da lista — `w:hidden` (sempre) ou `w:semiHidden` (até ser usado). */
  readonly hidden: boolean
  /** Criado por quem escreveu o documento, e não embutido no Word (`w:customStyle`). */
  readonly custom: boolean
  readonly basedOn?: string | undefined
  readonly next?: string | undefined
  /** O estilo de caractere ligado a este (`w:link`). */
  readonly link?: string | undefined
  readonly uiPriority?: number | undefined
  readonly paragraph?: StyleParagraphFormat | undefined
  readonly character?: StyleCharacterFormat | undefined
}

/**
 * O `w:docDefaults` e os estilos marcados `w:default="1"`.
 *
 * Os padrões são a base sobre a qual todo estilo se aplica; os ids são o que vale
 * para o parágrafo que não aponta estilo nenhum. Sem eles não há como responder
 * "qual é o estilo deste parágrafo" num documento que chama o estilo padrão de
 * `Padro` — e é essa a pergunta do painel.
 */
export interface StyleDefaults {
  readonly paragraph: StyleParagraphFormat
  readonly character: StyleCharacterFormat
  readonly paragraphStyleId: string | null
  readonly characterStyleId: string | null
}

export interface StyleSheet {
  readonly defaults: StyleDefaults
  /** Por id, que é o que o parágrafo aponta. */
  readonly styles: Readonly<Record<string, StyleDefinition>>
}

/**
 * A entrelinha 1,5 do CSS, em múltiplos da altura natural da linha.
 *
 * 1,5 ÷ 1,1499 (a altura da Liberation Serif, ver `line-metrics.ts`) = 1,3042. O
 * múltiplo do arquivo é medido sobre a altura da fonte, e não sobre o tamanho
 * dela; tratar um pelo outro encurtava cada linha em 15 %.
 *
 * Quatro casas, e não mais: o `w:line` vive numa grade de 240-avos (1,3042 × 240
 * = 313), e é nessa grade que o número volta do arquivo.
 */
export const BODY_LINE_FACTOR = 1.3042

const HEADING_SIZES = ['22pt', '17pt', '14pt', '12pt', '10pt', '8pt'] as const

/**
 * O antes do título é o `margin-top: 1em` do CSS (0,6em no `h5` e no `h6`, que
 * não têm regra própria), em pontos sobre o tamanho de cada um.
 */
const HEADING_BEFORE = [22, 17, 14, 12, 6, 4.8] as const

/**
 * O depois é a margem de baixo do navegador: 0,67em no `h1`, 0,83em no `h2`, 1em
 * no `h3`, 1,33em no `h4`, 1,67em e 2,33em nos dois últimos.
 */
const HEADING_AFTER = [14.75, 14.1, 14, 15.95, 16.65, 18.75] as const

/** "Manter com o próximo" é o `break-after: avoid` da impressão, que só os quatro primeiros têm. */
const HEADING_KEEP_NEXT = 4

function heading(level: number): StyleDefinition {
  const index = level - 1
  const paragraph: StyleParagraphFormat = {
    spaceBefore: HEADING_BEFORE[index]!,
    spaceAfter: HEADING_AFTER[index]!,
    ...(level <= HEADING_KEEP_NEXT ? { keepNext: true } : {}),
    outlineLevel: index,
  }

  return {
    id: `Heading${level}`,
    // O nome interno, em inglês e minúsculo, porque é assim que ele existe no
    // arquivo: é por ele que o leitor e o escritor reconhecem um título em
    // documento de qualquer idioma. Quem traduz para a tela é `styleLabelOf`.
    name: `heading ${level}`,
    type: StyleType.Paragraph,
    qFormat: true,
    hidden: false,
    custom: false,
    basedOn: 'Normal',
    next: 'Normal',
    uiPriority: 9,
    paragraph,
    character: { fontSize: HEADING_SIZES[index]!, bold: true },
  }
}

/**
 * Os estilos dos **arquivos antigos** — a tela de antes dos estilos, em dado.
 *
 * É a tabela que um `.sdoc` gravado antes da versão 3 do formato recebe ao ser
 * aberto, e a dos títulos que o escritor acrescenta a um DOCX que não os tem
 * (`BuiltinStyles.cs`, comparada com esta por
 * `src/main/sidecar/builtin-styles.test.ts`). Por isso ela não é um gosto: é o
 * CSS que o editor desenhava antes de desenhar a partir de estilos — Times New
 * Roman 12 pt, entrelinha 1,5, 0,6em antes e 1em depois — mais o padrão do
 * navegador para os títulos. Um número diferente daqui faz documento antigo
 * reabrir com outra paginação.
 *
 * O `#111111` do texto fica de fora de propósito: gravado como cor, ele voltaria
 * do arquivo como cor explícita em cada trecho reaberto.
 */
export const LEGACY_STYLES: StyleSheet = {
  defaults: {
    paragraph: {},
    // A fonte vai no padrão do documento, e não no `Normal`, porque é lá que o
    // Word a procura — e é de lá que todo estilo a herda.
    character: { fontFamily: 'Times New Roman', fontSize: '12pt' },
    paragraphStyleId: 'Normal',
    characterStyleId: 'DefaultParagraphFont',
  },
  styles: {
    Normal: {
      id: 'Normal',
      name: 'Normal',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
      paragraph: {
        // 0,6em antes (`.page__content > * + *`) e 1em depois (o padrão do
        // navegador para `p`), sobre 12 pt.
        spaceBefore: 7.2,
        spaceAfter: 12,
        lineSpacing: { kind: 'multiple', factor: BODY_LINE_FACTOR },
      },
    },
    DefaultParagraphFont: {
      id: 'DefaultParagraphFont',
      name: 'Default Paragraph Font',
      type: StyleType.Character,
      qFormat: false,
      hidden: true,
      custom: false,
      uiPriority: 1,
    },
    Heading1: heading(1),
    Heading2: heading(2),
    Heading3: heading(3),
    Heading4: heading(4),
    Heading5: heading(5),
    Heading6: heading(6),
    ListParagraph: {
      id: 'ListParagraph',
      name: 'List Paragraph',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
      basedOn: 'Normal',
      uiPriority: 34,
      // Meia polegada, que é o passo de recuo do OOXML (720 twips).
      paragraph: { indentMm: 12.7, contextualSpacing: true },
    },
    Hyperlink: {
      id: 'Hyperlink',
      name: 'Hyperlink',
      type: StyleType.Character,
      qFormat: false,
      hidden: false,
      custom: false,
      basedOn: 'DefaultParagraphFont',
      uiPriority: 99,
      character: { color: '#0563c1', underline: true },
    },
  },
}

/**
 * A entrelinha do Word 2013–2021, que o diálogo dele mostra como 1,08: 259 de
 * 240 avos, escrito como o leitor o devolve — a mesma grade de `BODY_LINE_FACTOR`.
 */
const WORD_LINE_FACTOR = 1.0792

/** Tamanho, antes e cor de cada título do Word 2013–2021. */
const WORD_HEADINGS = [
  { fontSize: '16pt', spaceBefore: 12, color: '#2f5496' },
  { fontSize: '13pt', spaceBefore: 2, color: '#2f5496' },
  { fontSize: '12pt', spaceBefore: 2, color: '#1f3763' },
  { spaceBefore: 2, color: '#2f5496', italic: true },
  { spaceBefore: 2, color: '#2f5496' },
  { spaceBefore: 2, color: '#1f3763' },
] as const

function wordHeading(level: number): StyleDefinition {
  const { spaceBefore, ...character } = WORD_HEADINGS[level - 1]!
  return {
    ...heading(level),
    // Sem a Calibri Light do Word: ela não tem substituta livre de mesmas
    // medidas, e um título medido com outra fonte quebra a linha noutro lugar.
    // O título herda a Calibri, que a Carlito desenha igual.
    paragraph: { spaceBefore, spaceAfter: 0, keepNext: true, keepLines: true, outlineLevel: level - 1 },
    character,
  }
}

/**
 * Os estilos do documento novo: o padrão do Word 2013–2021.
 *
 * Calibri 11 pt, entrelinha 1,08, 8 pt depois — no padrão do documento, e não no
 * `Normal`, porque é assim que o Word o grava e é de lá que todo estilo o
 * herda. É a tabela que o pacote DOCX do documento novo grava (`docx.create`).
 */
export const BUILTIN_STYLES: StyleSheet = {
  defaults: {
    paragraph: { spaceAfter: 8, lineSpacing: { kind: 'multiple', factor: WORD_LINE_FACTOR } },
    character: { fontFamily: 'Calibri', fontSize: '11pt' },
    paragraphStyleId: 'Normal',
    characterStyleId: 'DefaultParagraphFont',
  },
  styles: {
    ...LEGACY_STYLES.styles,
    Normal: {
      id: 'Normal',
      name: 'Normal',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
    },
    Heading1: wordHeading(1),
    Heading2: wordHeading(2),
    Heading3: wordHeading(3),
    Heading4: wordHeading(4),
    Heading5: wordHeading(5),
    Heading6: wordHeading(6),
  },
}

/**
 * O nome interno dos títulos, traduzido para a tela.
 *
 * Só estes: são os nomes que **nós** gravamos, em inglês, porque é a forma que o
 * arquivo exige para que o Word reconheça um título. Mostrar "heading 1" num
 * programa em português seria mostrar o arquivo, e não o documento. O nome de
 * qualquer outro estilo é dado do documento e aparece como está — traduzir o que
 * o autor escreveu seria inventar.
 */
const HEADING_LABELS: Readonly<Record<string, number>> = {
  'heading 1': 1,
  'heading 2': 2,
  'heading 3': 3,
  'heading 4': 4,
  'heading 5': 5,
  'heading 6': 6,
}

/** Como o estilo se chama na tela. */
export function styleLabelOf(style: StyleDefinition, language: Language = Language.Portuguese): string {
  const heading = HEADING_LABELS[style.name.toLowerCase()]
  return heading === undefined
    ? style.name
    : translate(language, 'document.styles.heading', { level: heading })
}

/**
 * Os estilos que o painel lista, na ordem em que ele os mostra.
 *
 * Os escondidos ficam de fora: `w:semiHidden` existe justamente para tirar da
 * lista o que é maquinaria do Word — `DefaultParagraphFont` e as dezenas de
 * variantes de tabela que todo documento do Word declara sem usar. Um painel com
 * elas é um painel que ninguém lê.
 *
 * A ordem é a do Word: primeiro a prioridade que o documento declara, e o nome
 * como critério de desempate — sem ele, dois estilos de mesma prioridade
 * trocariam de lugar entre uma abertura e outra.
 */
export function listedStyles(
  sheet: StyleSheet,
  language: Language = Language.Portuguese,
): readonly StyleDefinition[] {
  return Object.values(sheet.styles)
    .filter((style) => !style.hidden)
    .sort((left, right) => {
      const byPriority = (left.uiPriority ?? 100) - (right.uiPriority ?? 100)
      return byPriority !== 0
        ? byPriority
        : styleLabelOf(left, language).localeCompare(
            styleLabelOf(right, language),
            language === Language.Portuguese ? 'pt-BR' : 'en-US',
          )
    })
}

/** O bloco do cursor, no mínimo que responde "qual é o estilo dele". */
export interface BlockStyleQuery {
  /** O tipo do nó do editor: `paragraph`, `heading`, `codeBlock`… */
  readonly type: string
  /** O `w:pStyle` que o bloco trouxe do arquivo, quando trouxe. */
  readonly styleId?: string | null | undefined
  /** O nível, quando o bloco é um título. */
  readonly level?: number | null | undefined
}

/**
 * O estilo do bloco onde está o cursor — ou `null` quando não há resposta honesta.
 *
 * Três caminhos, e nessa ordem:
 *
 * 1. o `w:pStyle` que o bloco trouxe, **se o documento define esse estilo**. Um
 *    id que o documento não define é silêncio no Word, e seria mentira aqui;
 * 2. o título, pelo **nome interno** do estilo (`heading 3`), que é o critério do
 *    leitor e do escritor (`StyleResolver.HeadingLevelByName`). Pelo id não
 *    funcionaria: num documento em alemão o estilo se chama `berschrift3`;
 * 3. o estilo padrão de parágrafo, que é o que vale sem `w:pStyle`.
 */
export function blockStyleOf(sheet: StyleSheet, block: BlockStyleQuery): StyleDefinition | null {
  const declared = block.styleId ?? null
  if (declared !== null && declared in sheet.styles) return sheet.styles[declared] ?? null

  if (block.type === 'heading' && typeof block.level === 'number') {
    const name = `heading ${block.level}`
    const found = Object.values(sheet.styles).find(
      (style) => style.type === StyleType.Paragraph && style.name.toLowerCase() === name,
    )
    if (found !== undefined) return found
  }

  const fallback = sheet.defaults.paragraphStyleId
  return (fallback !== null ? sheet.styles[fallback] : undefined) ?? null
}
