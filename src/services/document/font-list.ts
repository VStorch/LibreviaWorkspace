/** List order and parsing of what each system returns; `src/main/system-fonts.ts` asks. */

/** The only ones with a metric-compatible substitute on any machine (see `fonts.ts`). */
export const GUARANTEED_FONT_FAMILIES: readonly string[] = [
  'Calibri',
  'Cambria',
  'Arial',
  'Times New Roman',
  'Courier New',
]

const key = (family: string): string => family.trim().toLowerCase()

/**
 * The document's, the guaranteed ones and the installed ones alphabetically. An empty installed
 * list is a system without `fontconfig`, not an error.
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

  // `fc-list` returns fontconfig cache order.
  for (const family of [...installed].sort((left, right) => left.localeCompare(right, 'pt-BR'))) {
    push(family)
  }

  return ordered
}

/** The picker knows names, and the document carries CSS stacks. */
export function firstFamilyOf(stack: string): string {
  const first = stack.split(',')[0]?.trim() ?? ''
  return first.replace(/^['"]|['"]$/g, '')
}

/** The font lives in the text mark and in the block attribute, and the reader emits both. */
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

/** One family per line; comma-separated names ("Nimbus Sans,Nimbus Sans L") keep the first. */
export function parseFontconfigFamilies(output: string): string[] {
  const families = new Set<string>()

  for (const line of output.split('\n')) {
    const family = line.split(',')[0]?.trim() ?? ''
    if (family.length > 0) families.add(family)
  }

  return [...families]
}

/**
 * The styles Windows glues to the name. Only these, and only at the end: "Arial Black" is a family.
 */
const WINDOWS_STYLE_SUFFIXES = [' Bold Italic', ' Bold Oblique', ' Bold', ' Italic', ' Oblique', ' Regular']

/** Each line is `    <names> (TrueType)    REG_SZ    file.ttf`, with styles separated by ` & `. */
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
      // Only when a name remains: "Bold" alone is the family name, odd as it is, and cutting it
      // would leave nothing.
      if (base.length > 0) return base
    }
  }

  return name
}
