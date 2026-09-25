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
 * Os estilos dos arquivos antigos, com os dois lados de verdade.
 *
 * A tabela existe duas vezes, e tem de existir: o sidecar acrescenta os títulos
 * dela ao DOCX que não os tem (`BuiltinStyles.cs`) e o leitor do `.sdoc` a dá a
 * todo arquivo gravado antes da versão 3 do formato (`LEGACY_STYLES`). Uma medida mudada de um
 * lado só não quebra nada visível na hora — e é justamente aí que está o perigo:
 * o documento passa a abrir com uma aparência e a ser gravado com outra, e a
 * paginação muda no arquivo de alguém que não editou nada.
 *
 * Mesmo espírito de `line-metrics.test.ts`: o C# é lido como texto, porque o que
 * precisa ser comparado são os **números**, e não o comportamento.
 *
 * Fica em `src/main` porque só aqui há Node: `src/services` é compilado também
 * para a web, e lá não existe `node:fs` para ler o arquivo do sidecar.
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
    // A fonte e o tamanho moram no `w:docDefaults`, e não no `Normal`: é de lá
    // que todo estilo os herda, e é lá que o Word os procura.
    expect(LEGACY_STYLES.defaults.character).toEqual({
      fontFamily: constantText(source, 'BodyFont'),
      fontSize: `${constantNumber(source, 'BodySizePt')}pt`,
    })
    expect(LEGACY_STYLES.defaults.paragraph).toEqual({})

    // O `w:default="1"` de cada tipo: é ele que responde qual estilo vale num
    // parágrafo sem `w:pStyle`.
    expect(LEGACY_STYLES.defaults.paragraphStyleId).toBe(defaultIdOf(source, false))
    expect(LEGACY_STYLES.defaults.characterStyleId).toBe(defaultIdOf(source, true))
  })

  it('a entrelinha do corpo é o mesmo fator nos dois lados', () => {
    // O número que o `w:line` recebe nasce daqui. Divergente, o documento novo
    // sairia do editor com uma entrelinha e voltaria do arquivo com outra.
    expect(constantNumber(source, 'BodyLineFactor')).toBe(BODY_LINE_FACTOR)
  })
})

/** O valor de uma `const` de texto do C#. */
function constantText(source: string, name: string): string {
  const match = new RegExp(`const string ${name} = "([^"]*)"`).exec(source)
  if (match?.[1] === undefined) throw new Error(`constante ${name} não encontrada`)
  return match[1]
}

/** O valor de uma `const` numérica do C#. */
function constantNumber(source: string, name: string): number {
  const match = new RegExp(`const double ${name} = ([\\d.]+)`).exec(source)
  if (match?.[1] === undefined) throw new Error(`constante ${name} não encontrada`)
  return Number(match[1])
}

/** O id do estilo marcado `Default: true` — de caractere ou de parágrafo. */
function defaultIdOf(source: string, character: boolean): string | null {
  const found = entriesOf(source)
    .map(argumentsOf)
    .find((args) => args['Default'] === true && (args['Character'] === true) === character)
  return typeof found?.['Id'] === 'string' ? found['Id'] : null
}

/**
 * Cada `new(...)` da tabela, como texto.
 *
 * Por varredura de parênteses, e não por expressão regular: as entradas ocupam
 * mais de uma linha, e uma expressão que casasse com a primeira `)` cortaria a
 * entrada no meio sem reclamar — um teste que compara metade dos dados é pior do
 * que nenhum.
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

/**
 * Os argumentos de uma entrada, por nome.
 *
 * Os dois primeiros são posicionais — o id e o nome, na ordem do construtor —, e
 * o resto vem nomeado. É a forma em que a tabela do C# está escrita, e ela é
 * legível justamente por isso.
 */
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
  // A entrelinha do corpo é declarada como constante no próprio arquivo, e é
  // comparada em teste próprio: aqui vale o número que ela carrega.
  if (text === 'BodyLineFactor') return BODY_LINE_FACTOR
  const number = Number(text)
  if (Number.isNaN(number)) throw new Error(`valor que o teste não sabe ler: ${text}`)
  return number
}

/** A tabela do C# na forma do modelo do documento, pronta para comparar. */
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
      // Os dois jeitos de esconder um estilo da galeria contam como um só no
      // modelo: a pergunta que o painel faz é "isto aparece na lista?".
      hidden: args['Hidden'] === true || args['SemiHidden'] === true,
      // A tabela é a dos estilos embutidos do Word; nenhum deles é criado por
      // quem escreveu o documento.
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

/** Meios-pontos não entram aqui: a tabela já está em pontos, como o editor os usa. */
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
