/** A ordem da lista e a leitura do que cada sistema devolve; quem pergunta é `src/main/system-fonts.ts`. */

/** As únicas com substituta metricamente compatível em qualquer máquina (ver `fonts.ts`). */
export const GUARANTEED_FONT_FAMILIES: readonly string[] = [
  'Calibri',
  'Cambria',
  'Arial',
  'Times New Roman',
  'Courier New',
]

const key = (family: string): string => family.trim().toLowerCase()

/**
 * As do documento, as garantidas e as instaladas em ordem alfabética. Lista de
 * instaladas vazia é o sistema sem `fontconfig`, e não erro.
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

  // `fc-list` devolve na ordem do cache do fontconfig.
  for (const family of [...installed].sort((left, right) => left.localeCompare(right, 'pt-BR'))) {
    push(family)
  }

  return ordered
}

/** O seletor conhece nomes, e o documento traz pilhas de CSS. */
export function firstFamilyOf(stack: string): string {
  const first = stack.split(',')[0]?.trim() ?? ''
  return first.replace(/^['"]|['"]$/g, '')
}

/** A fonte mora na marca do texto e no atributo do bloco, e o leitor a emite nos dois. */
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

/** Uma família por linha; nomes separados por vírgula ("Nimbus Sans,Nimbus Sans L") ficam no primeiro. */
export function parseFontconfigFamilies(output: string): string[] {
  const families = new Set<string>()

  for (const line of output.split('\n')) {
    const family = line.split(',')[0]?.trim() ?? ''
    if (family.length > 0) families.add(family)
  }

  return [...families]
}

/** Os cortes que o Windows cola no nome. Só estes, e só no fim: "Arial Black" é família. */
const WINDOWS_STYLE_SUFFIXES = [' Bold Italic', ' Bold Oblique', ' Bold', ' Italic', ' Oblique', ' Regular']

/** Cada linha é `    <nomes> (TrueType)    REG_SZ    arquivo.ttf`, com cortes separados por ` & `. */
export function parseWindowsFontRegistry(output: string): string[] {
  const families = new Set<string>()

  for (const line of output.split('\n')) {
    const match = /^\s+(.+?)\s{2,}REG_SZ\s{2,}/.exec(line)
    if (match === null) continue

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
