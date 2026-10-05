import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  BODY_LINE_FACTOR,
  LEGACY_STYLES,
  type StyleCharacterFormat,
  type StyleDefinition,
  type StyleParagraphFormat,
} from '@services/document/styles.js'

/**
 * The styles of old files, with both real sides: the sidecar adds headings to a DOCX that lacks
 * them (`BuiltinStyles.cs`) and the `.sdoc` reader gives them to every file before version 3
 * (`LEGACY_STYLES`). A value changed on one side only changes the pagination of someone who edited
 * nothing.
 *
 * The C# is read as text, because what is compared are numbers. Lives in `src/main`, the only place
 * with `node:fs`.
 */
describe('contrato dos estilos dos arquivos antigos', () => {
  const source = readFileSync(
    new URL('../../../sidecar/src/Librevia.Format/Docx/BuiltinStyles.cs', import.meta.url),
    'utf8',
  )

  it('a tabela do sidecar é a mesma do modelo do documento', () => {
    const declared = parseTable(source)

    expect(Object.keys(declared).length).toBeGreaterThan(0)
    expect(declared).toEqual(LEGACY_STYLES.styles)
  })

  it('os padrões do documento são os mesmos', () => {
    // Font and size live in `w:docDefaults`, which every style inherits from.
    expect(LEGACY_STYLES.defaults.character).toEqual({
      fontFamily: constantText(source, 'BodyFont'),
      fontSize: `${constantNumber(source, 'BodySizePt')}pt`,
    })
    expect(LEGACY_STYLES.defaults.paragraph).toEqual({})

    // `w:default="1"` says which style applies without `w:pStyle`.
    expect(LEGACY_STYLES.defaults.paragraphStyleId).toBe(defaultIdOf(source, false))
    expect(LEGACY_STYLES.defaults.characterStyleId).toBe(defaultIdOf(source, true))
  })

  it('a entrelinha do corpo é o mesmo fator nos dois lados', () => {
    // If they differ, the document would leave the editor with one line spacing and come back with
    // another.
    expect(constantNumber(source, 'BodyLineFactor')).toBe(BODY_LINE_FACTOR)
  })
})

function constantText(source: string, name: string): string {
  const match = new RegExp(`const string ${name} = "([^"]*)"`).exec(source)
  if (match?.[1] === undefined) throw new Error(`constante ${name} não encontrada`)
  return match[1]
}

function constantNumber(source: string, name: string): number {
  const match = new RegExp(`const double ${name} = ([\\d.]+)`).exec(source)
  if (match?.[1] === undefined) throw new Error(`constante ${name} não encontrada`)
  return Number(match[1])
}

/** Character or paragraph. */
function defaultIdOf(source: string, character: boolean): string | null {
  const found = entriesOf(source)
    .map(argumentsOf)
    .find((args) => args['Default'] === true && (args['Character'] === true) === character)
  return typeof found?.['Id'] === 'string' ? found['Id'] : null
}

/**
 * Scans parentheses: entries span several lines, and a regular expression would stop at the first
 * `)` without complaint.
 */
function entriesOf(source: string): string[] {
  const start = source.indexOf('BuiltinStyle[] All =')
  if (start < 0) throw new Error('tabela All não encontrada')

  const entries: string[] = []
  for (let at = source.indexOf('new(', start); at >= 0; at = source.indexOf('new(', at + 1)) {
    let depth = 0
    let quoted = false
    for (let index = at + 3; index < source.length; index += 1) {
      const character = source[index]
      if (quoted) {
        if (character === '"') quoted = false
        continue
      }
      if (character === '"') quoted = true
      else if (character === '(') depth += 1
      else if (character === ')') {
        depth -= 1
        if (depth === 0) {
          entries.push(source.slice(at + 4, index))
          at = index
          break
        }
      }
    }
  }

  if (entries.length === 0) throw new Error('nenhuma entrada na tabela All')
  return entries
}

type Argument = string | number | boolean

/** Id and name positional, the rest named. */
function argumentsOf(entry: string): Record<string, Argument> {
  const parts: string[] = []
  let current = ''
  let quoted = false
  for (const character of entry) {
    if (quoted) {
      current += character
      if (character === '"') quoted = false
      continue
    }
    if (character === '"') quoted = true
    if (character === ',') {
      parts.push(current)
      current = ''
      continue
    }
    current += character
  }
  parts.push(current)

  const positional = ['Id', 'Name']
  const args: Record<string, Argument> = {}
  for (const part of parts) {
    const text = part.trim()
    if (text.length === 0) continue

    const named = /^(\w+):\s*(.+)$/.exec(text)
    if (named?.[1] !== undefined && named[2] !== undefined) {
      args[named[1]] = valueOf(named[2].trim())
      continue
    }

    const name = positional.shift()
    if (name === undefined) throw new Error(`argumento posicional a mais em: ${entry}`)
    args[name] = valueOf(text)
  }

  return args
}

function valueOf(text: string): Argument {
  if (text.startsWith('"')) return text.slice(1, -1)
  if (text === 'true') return true
  if (text === 'false') return false
  // The body line spacing is a file constant, compared in its own test.
  if (text === 'BodyLineFactor') return BODY_LINE_FACTOR
  const number = Number(text)
  if (Number.isNaN(number)) throw new Error(`valor que o teste não sabe ler: ${text}`)
  return number
}

function parseTable(source: string): Record<string, StyleDefinition> {
  const styles: Record<string, StyleDefinition> = {}

  for (const entry of entriesOf(source)) {
    const args = argumentsOf(entry)
    const id = String(args['Id'])

    const paragraph: StyleParagraphFormat = {
      ...number(args, 'BeforePt', 'spaceBefore'),
      ...number(args, 'AfterPt', 'spaceAfter'),
      ...(typeof args['LineFactor'] === 'number'
        ? { lineSpacing: { kind: 'multiple' as const, factor: args['LineFactor'] } }
        : {}),
      ...number(args, 'IndentMm', 'indentMm'),
      ...flag(args, 'KeepNext', 'keepNext'),
      ...flag(args, 'ContextualSpacing', 'contextualSpacing'),
      ...number(args, 'OutlineLevel', 'outlineLevel'),
    }

    const character: StyleCharacterFormat = {
      ...(typeof args['SizePt'] === 'number' ? { fontSize: pointsText(args['SizePt']) } : {}),
      ...flag(args, 'Bold', 'bold'),
      ...(typeof args['Color'] === 'string' ? { color: args['Color'] } : {}),
      ...flag(args, 'Underline', 'underline'),
    }

    styles[id] = {
      id,
      name: String(args['Name']),
      type: args['Character'] === true ? 'character' : 'paragraph',
      qFormat: args['QFormat'] === true,
      // The two ways of hiding from the gallery are one in the model.
      hidden: args['Hidden'] === true || args['SemiHidden'] === true,
      // Word's builtin styles.
      custom: false,
      ...text(args, 'BasedOn', 'basedOn'),
      ...text(args, 'Next', 'next'),
      ...text(args, 'Link', 'link'),
      ...number(args, 'UiPriority', 'uiPriority'),
      ...(Object.keys(paragraph).length > 0 ? { paragraph } : {}),
      ...(Object.keys(character).length > 0 ? { character } : {}),
    }
  }

  return styles
}

/** No half-points here: the table is already in points, as the editor uses them. */
function pointsText(points: number): string {
  return `${points}pt`
}

function number(args: Record<string, Argument>, from: string, to: string): Record<string, number> {
  const value = args[from]
  return typeof value === 'number' ? { [to]: value } : {}
}

function text(args: Record<string, Argument>, from: string, to: string): Record<string, string> {
  const value = args[from]
  return typeof value === 'string' ? { [to]: value } : {}
}

function flag(args: Record<string, Argument>, from: string, to: string): Record<string, true> {
  return args[from] === true ? { [to]: true } : {}
}
