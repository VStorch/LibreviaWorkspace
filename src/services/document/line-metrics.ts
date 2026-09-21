/**
 * A altura natural da linha de cada fonte, e as duas contas que dependem dela.
 *
 * O múltiplo do OOXML (`w:line` em 240-avos, com `w:lineRule="auto"`) é medido
 * sobre a **altura natural da fonte** — `(ascender - descender + lineGap) /
 * unitsPerEm` da tabela `hhea` —, e não sobre o tamanho dela. É a conta que o
 * Word e o LibreOffice fazem, e a mesma que o navegador chama de
 * `line-height: normal`.
 *
 * Por isso o atributo `lineHeight` do modelo **não** guarda o fator do Word:
 * guarda o número que o CSS entende, que é o fator já multiplicado pela altura
 * natural. O leitor (`BodyReader.LineHeightOf`) multiplica na entrada e o
 * gravador (`ParagraphFormat.ApplyLineHeight`) divide na saída, e as duas pontas
 * fecham. Quem não fechava era a interface: o diálogo de parágrafo tratava o
 * número do CSS como se fosse o do Word, então todo parágrafo de Calibri abria
 * como "Múltiplo 1,1499" e escolher "1,5" gravava 1,23 linha no arquivo — erro
 * de 15 a 22 %, dependendo da fonte.
 *
 * A tabela mora aqui, e não só no C#, porque agora as duas conversões acontecem
 * nos dois lados: `line-metrics.test.ts` compara esta tabela com `LineMetrics.cs`
 * a cada execução, para que uma fonte acrescentada lá não fique faltando aqui.
 *
 * Só as fontes que o instalador leva, e as que elas substituem. Para o resto não
 * há palpite honesto — a substituta depende da máquina.
 */

/** A fonte do editor quando o documento não diz outra. */
export const DEFAULT_NATURAL_LINE_HEIGHT = 1.1499

const LIBERATION_SERIF = DEFAULT_NATURAL_LINE_HEIGHT

/** Nome da família → altura natural. As chaves comparam sem caixa nem espaço. */
export const NATURAL_LINE_HEIGHTS: Readonly<Record<string, number>> = {
  arial: 1.1499,
  helvetica: 1.1499,
  'liberation sans': 1.1499,
  'times new roman': LIBERATION_SERIF,
  'liberation serif': LIBERATION_SERIF,
  'courier new': 1.1328,
  'liberation mono': 1.1328,
  calibri: 1.2207,
  carlito: 1.2207,
  cambria: 1.15,
  caladea: 1.15,
}

/**
 * A primeira família da pilha do CSS.
 *
 * O modelo guarda `"Calibri, sans-serif"`, porque é o que o leitor monta a partir
 * da tabela de fontes do arquivo. A altura é da primeira — a genérica existe para
 * o caso de a primeira faltar, e aí ninguém sabe qual arquivo o navegador usa.
 * É o espelho de `ParagraphFormat.FirstFont`.
 */
export function firstFontOf(stack: string | null | undefined): string | null {
  if (typeof stack !== 'string') return null
  const first =
    stack
      .split(',')[0]
      ?.trim()
      .replace(/^['"]|['"]$/g, '') ?? ''
  return first.length > 0 ? first : null
}

/**
 * A altura natural da fonte, ou `null` quando não se sabe qual arquivo o
 * navegador vai usar. É o espelho de `LineMetrics.Of`: pilha vazia devolve o
 * padrão do editor, e não `null` — sem fonte declarada, a fonte é a nossa.
 */
export function naturalLineHeightOf(stack: string | null | undefined): number | null {
  const first = firstFontOf(stack)
  if (first === null) return DEFAULT_NATURAL_LINE_HEIGHT
  return NATURAL_LINE_HEIGHTS[first.toLowerCase()] ?? null
}

/**
 * O fator do Word → o número que o CSS (e o atributo) entendem.
 *
 * Espelha `BodyReader.Multiple`, inclusive no caso da fonte desconhecida: aí o
 * fator 1 sai como `normal`, que é o silêncio do arquivo, e o resto sai pelo
 * palpite de 1,15 — a altura de quase toda fonte latina, e a mesma que o
 * gravador usa para desfazer a conta.
 */
export function cssLineHeightOf(factor: number, stack: string | null | undefined): string {
  const natural = naturalLineHeightOf(stack)
  if (natural !== null) return numberText(factor * natural)
  return factor === 1 ? 'normal' : numberText(factor * DEFAULT_NATURAL_LINE_HEIGHT)
}

/**
 * O mesmo, mas nunca `normal`.
 *
 * Para quando é preciso **sobrescrever** uma entrelinha declarada: `normal` é o
 * único valor que o gravador não grava — ele quer dizer "não mexa no `w:line`" —,
 * e aí a medida antiga ficaria de pé no arquivo sem ninguém avisar.
 */
export function explicitCssLineHeightOf(factor: number, stack: string | null | undefined): string {
  return numberText(factor * (naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT))
}

/**
 * O número do CSS → o fator do Word, que é o que o diálogo mostra.
 *
 * Arredondado ao centésimo porque é a precisão que sobrevive à ida e volta: o
 * `w:line` vem em 240-avos (um passo de 0,004) e o atributo guarda quatro casas.
 * Sem o arredondamento, "1,5" voltaria como 1,4999 e o seletor da barra perderia
 * a opção.
 */
export function lineFactorOf(css: number, stack: string | null | undefined): number {
  const natural = naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT
  return Math.round((css / natural) * 100) / 100
}

/** Quatro casas no máximo, sem zero à direita — a forma que o leitor escreve. */
function numberText(value: number): string {
  return String(Math.round(value * 10_000) / 10_000)
}
