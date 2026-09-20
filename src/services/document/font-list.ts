/**
 * A lista de fontes que a barra oferece.
 *
 * Eram sete nomes fixos no código, e isso mentia duas vez: escondia as fontes
 * que a máquina tem e escondia as que o **documento aberto** pede. A segunda é a
 * pior — um documento em Garamond abria com "Fonte padrão" no seletor, e trocar
 * qualquer outra coisa na barra o reescrevia noutra fonte sem aviso.
 *
 * Aqui mora a parte que não depende de sistema operacional nenhum: a ordem da
 * lista e a leitura do que cada sistema cospe quando lhe perguntam as fontes.
 * Quem pergunta é o processo main (`src/main/system-fonts.ts`), porque só ele
 * pode executar programa.
 */

/**
 * As famílias que viajam no instalador.
 *
 * Vêm primeiro depois das do documento porque são as únicas com garantia: existe
 * substituta metricamente compatível para cada uma (ver `fonts.ts`), então
 * escolhê-las dá o mesmo resultado em qualquer máquina. As demais dependem do
 * que estiver instalado.
 */
export const GUARANTEED_FONT_FAMILIES: readonly string[] = [
  'Calibri',
  'Cambria',
  'Arial',
  'Times New Roman',
  'Courier New',
]

/** Normaliza para comparar: fonte é nome, e nome não distingue caixa. */
const key = (family: string): string => family.trim().toLowerCase()

/**
 * A lista final, sem repetição e na ordem em que a pessoa procura.
 *
 * Três camadas, e a ordem delas é a resposta a "onde está a minha fonte?":
 *
 *  1. **as do documento** — quem abriu um arquivo alheio quer ver o que ele usa;
 *  2. **as garantidas** — as cinco que o instalador leva, que funcionam sempre;
 *  3. **as instaladas**, em ordem alfabética da região.
 *
 * Lista de instaladas vazia não é caso de erro: é o que acontece num sistema sem
 * `fontconfig`, e aí sobram as duas primeiras camadas — exatamente o que a barra
 * oferecia antes.
 */
export function orderFontFamilies(
  installed: readonly string[],
  inDocument: readonly string[] = [],
): string[] {
  const seen = new Set<string>()
  const ordered: string[] = []

  const push = (family: string): void => {
    const name = family.trim()
    if (name.length === 0 || seen.has(key(name))) return
    seen.add(key(name))
    ordered.push(name)
  }

  for (const family of inDocument) push(family)
  for (const family of GUARANTEED_FONT_FAMILIES) push(family)

  // Ordenada aqui, e não por quem coletou: `fc-list` devolve na ordem do cache
  // do fontconfig, que não é ordem nenhuma para quem lê.
  for (const family of [...installed].sort((left, right) => left.localeCompare(right, 'pt-BR'))) {
    push(family)
  }

  return ordered
}

/**
 * O primeiro nome de uma pilha de CSS.
 *
 * O que o documento traz em `fontFamily` é uma pilha — a fonte pedida e a
 * substituta genérica atrás — e o seletor conhece nomes, não pilhas.
 */
export function firstFamilyOf(stack: string): string {
  const first = stack.split(',')[0]?.trim() ?? ''
  return first.replace(/^['"]|['"]$/g, '')
}

/**
 * As famílias que aparecem no documento, na ordem em que aparecem.
 *
 * Varre o modelo inteiro porque a fonte mora em dois lugares: na marca do texto
 * (`textStyle`) e no atributo do bloco — a altura da linha nasce da fonte do
 * elemento, e por isso o leitor a emite nos dois. Sem os dois, um título em
 * Garamond não entraria na lista.
 */
export function familiesInDocument(doc: unknown): string[] {
  const found: string[] = []
  const seen = new Set<string>()

  const note = (value: unknown): void => {
    if (typeof value !== 'string') return
    const family = firstFamilyOf(value)
    if (family.length === 0 || seen.has(key(family))) return
    seen.add(key(family))
    found.push(family)
  }

  const walk = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) walk(child)
      return
    }
    if (node === null || typeof node !== 'object') return

    const record = node as Record<string, unknown>
    const attrs = record['attrs']
    if (attrs !== null && typeof attrs === 'object') {
      note((attrs as Record<string, unknown>)['fontFamily'])
    }

    walk(record['marks'])
    walk(record['content'])
  }

  walk(doc)
  return found
}

/**
 * A saída de `fc-list --format '%{family[0]}\n'`.
 *
 * Uma família por linha, e o fontconfig às vezes devolve mais de um nome para a
 * mesma família separados por vírgula — o nome local e o inglês, como em
 * "Nimbus Sans,Nimbus Sans L". O primeiro basta: é o que o CSS acha.
 */
export function parseFontconfigFamilies(output: string): string[] {
  const families = new Set<string>()

  for (const line of output.split('\n')) {
    const family = line.split(',')[0]?.trim() ?? ''
    if (family.length > 0) families.add(family)
  }

  return [...families]
}

/**
 * Os sufixos de corte que o Windows cola no nome do arquivo de fonte.
 *
 * "Arial Bold" não é família: é o corte gordo da Arial, e listá-lo faria o
 * seletor oferecer duas Arial. Só estes cinco, e sempre no fim do nome: cortar
 * mais apagaria famílias de verdade — "Arial Black" e "Segoe UI Light" existem
 * como famílias próprias.
 */
const WINDOWS_STYLE_SUFFIXES = [' Bold Italic', ' Bold Oblique', ' Bold', ' Italic', ' Oblique', ' Regular']

/**
 * A saída de `reg query` sobre a chave de fontes do Windows.
 *
 * Cada linha é `    <nomes> (TrueType)    REG_SZ    arquivo.ttf`, e `<nomes>`
 * pode trazer vários cortes de uma vez, separados por ` & ` — é assim que o
 * instalador de fonte registra uma família inteira num arquivo só.
 */
export function parseWindowsFontRegistry(output: string): string[] {
  const families = new Set<string>()

  for (const line of output.split('\n')) {
    const match = /^\s+(.+?)\s{2,}REG_SZ\s{2,}/.exec(line)
    if (match === null) continue

    // O tipo do arquivo vem entre parênteses no fim, e não faz parte do nome.
    const names = match[1]!.replace(/\s*\((TrueType|OpenType|All res)\)\s*$/i, '')

    for (const name of names.split('&')) {
      const family = stripStyleSuffix(name.trim())
      if (family.length > 0) families.add(family)
    }
  }

  return [...families]
}

function stripStyleSuffix(name: string): string {
  for (const suffix of WINDOWS_STYLE_SUFFIXES) {
    if (name.toLowerCase().endsWith(suffix.toLowerCase())) {
      const base = name.slice(0, -suffix.length).trim()
      // Só quando sobra nome: "Bold" sozinho é o nome da família, por estranho
      // que pareça, e cortá-lo devolveria vazio.
      if (base.length > 0) return base
    }
  }

  return name
}
