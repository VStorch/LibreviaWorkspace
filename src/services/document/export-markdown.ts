import {
  imageData,
  imageExtension,
  isSectionMarkOnly,
  prepareExport,
  safeHref,
  tocLevelOf,
  withoutPageNumbers,
  type ExportNote,
  type ExportSource,
} from './export-common.js'
import { createHtmlRenderer, escapeHtml, mathHtml } from './export-html.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { latexOfEquation } from './mathml-latex.js'

/**
 * Exportação para Markdown: CommonMark com as tabelas e as notas do GFM.
 *
 * As imagens vão para uma **pasta irmã** (`relatorio_arquivos/`), referidas por
 * caminho relativo — como o "Salvar como página da Web" do Word. Embuti-las em
 * `data:` deixaria o texto ilegível no editor de texto, que é onde Markdown se
 * lê, e o GitHub e boa parte dos visualizadores recusam imagem em `data:`. Esta
 * função só devolve as imagens; quem grava é o processo main.
 *
 * O que o Markdown não tem vira o mais próximo que ele tem: a tabela com
 * células mescladas sai em HTML (que o CommonMark aceita no meio do texto), o
 * sobrescrito em `<sup>`, as notas de fim também em `[^n]`. Sublinhado, cor e
 * fonte se perdem — o Markdown não as conhece.
 */

export interface MarkdownExportOptions {
  /** O nome da pasta das imagens, ao lado do arquivo (`relatorio_arquivos`). */
  readonly assetFolder: string
}

/** Uma imagem a gravar na pasta, com o nome que o texto usa. */
export interface MarkdownAsset {
  readonly name: string
  readonly mime: string
  readonly base64: string
}

export interface MarkdownExport {
  readonly markdown: string
  readonly assets: readonly MarkdownAsset[]
}

type Mode = 'block' | 'cell' | 'heading'

interface OpenMark {
  readonly key: string
  readonly open: string
  readonly close: string
  /** Onde o conteúdo da marca começa na saída. */
  readonly at: number
}

/** A ordem de abertura: o link por fora, o sobrescrito por dentro. */
const MARK_ORDER = ['link', 'bold', 'italic', 'strike', 'superscript', 'subscript']

export function exportMarkdown(
  model: Pick<DocumentModel, 'doc' | 'notes'>,
  options: MarkdownExportOptions,
): MarkdownExport {
  const source = prepareExport(model)
  const assets: MarkdownAsset[] = []
  const assetBySrc = new Map<string, string>()

  /** O caminho relativo da imagem, gravando-a na lista uma vez só. */
  function imagePath(src: unknown): string | null {
    const data = imageData(src)
    if (data === null) return null
    const key = String(src)
    let name = assetBySrc.get(key)
    if (name === undefined) {
      name = `image-${assets.length + 1}.${imageExtension(data.mime)}`
      assets.push({ name, mime: data.mime, base64: data.base64 })
      assetBySrc.set(key, name)
    }
    return `${encodePathPart(options.assetFolder)}/${encodePathPart(name)}`
  }

  const writer = new MarkdownWriter(source, imagePath)
  const body = writer.blocks(source.doc.content ?? [])
  const notes = source.notes.map((note) => writer.noteDefinition(note)).join('\n\n')
  const markdown = [body, notes].filter((part) => part !== '').join('\n\n')
  return { markdown: markdown === '' ? '' : `${markdown}\n`, assets }
}

class MarkdownWriter {
  private readonly html

  constructor(
    private readonly source: ExportSource,
    private readonly imagePath: (src: unknown) => string | null,
  ) {
    this.html = createHtmlRenderer(source, imagePath)
  }

  blocks(nodes: readonly DocumentNode[]): string {
    const parts: string[] = []
    let previous: DocumentNode | null = null
    for (const node of nodes) {
      const text = this.block(node)
      if (text === '') continue
      // Duas listas do mesmo tipo seguidas virariam uma só.
      if (previous !== null && previous.type === node.type && isList(node)) parts.push('<!-- -->')
      parts.push(text)
      previous = node
    }
    return parts.join('\n\n')
  }

  noteDefinition(note: ExportNote): string {
    const body = this.blocks(note.body)
    return `[^${noteKey(note)}]: ${indentRest(body, '    ')}`
  }

  private block(node: DocumentNode): string {
    switch (node.type) {
      case 'paragraph':
        return isSectionMarkOnly(node) ? '' : this.paragraph(node.content ?? [])
      case 'heading': {
        const level = Math.min(6, Math.max(1, Number(node.attrs?.['level']) || 1))
        const text = this.inline(node.content ?? [], 'heading').trim()
        if (text === '') return ''
        // O `#` no fim seria lido como o fecho opcional do título.
        return `${'#'.repeat(level)} ${text.replace(/#$/, '\\#')}`
      }
      case 'bulletList':
      case 'orderedList':
        return this.list(node)
      case 'table':
        return this.table(node)
      case 'blockquote': {
        const inner = this.blocks(node.content ?? [])
        return inner === '' ? '' : prefixLines(inner, '> ', '>')
      }
      case 'codeBlock':
        return codeBlock(node)
      case 'horizontalRule':
        return '---'
      case 'pageBreak':
        return ''
      case 'tableOfContents':
        return this.contents(node)
      default:
        return node.content === undefined ? '' : this.blocks(node.content)
    }
  }

  private paragraph(content: readonly DocumentNode[]): string {
    const text = this.inline(content, 'block')
      .replace(/(\\\n)+$/, '')
      .replace(/^\n+|\n+$/g, '')
    return text.trim() === '' ? '' : escapeLineStarts(text)
  }

  private list(node: DocumentNode): string {
    const ordered = node.type === 'orderedList'
    const items = (node.content ?? []).filter((child) => child.type === 'listItem')
    let next = items[0] === undefined ? 1 : (this.source.itemOf.get(items[0])?.value ?? 1)
    const loose = items.some((item) => (item.content ?? []).filter((child) => !isList(child)).length > 1)

    const rendered = items.map((item) => {
      const value = this.source.itemOf.get(item)?.value ?? next
      next = value + 1
      const marker = ordered ? `${value}.` : '-'
      const pad = ' '.repeat(marker.length + 1)
      const children = item.content ?? []
      const parts: string[] = []
      children.forEach((child, index) => {
        const text = this.block(child)
        if (text === '') return
        // A sublista cola no parágrafo do item: lista compacta continua compacta.
        const glue = parts.length === 0 ? '' : isList(child) && !loose && index > 0 ? '\n' : '\n\n'
        parts.push(glue + text)
      })
      const body = parts.join('')
      return body === '' ? marker : `${marker} ${indentRest(body, pad)}`
    })
    return rendered.join(loose ? '\n\n' : '\n')
  }

  private table(node: DocumentNode): string {
    const rows = (node.content ?? []).filter((row) => row.type === 'tableRow')
    if (rows.length === 0) return ''
    if (!isSimpleTable(rows)) return this.html.blocks([node])

    const lines = rows.map((row) => `| ${(row.content ?? []).map((cell) => this.cell(cell)).join(' | ')} |`)
    const columns = (rows[0]!.content ?? []).length
    lines.splice(1, 0, `| ${Array.from({ length: columns }, () => '---').join(' | ')} |`)
    return lines.join('\n')
  }

  private cell(node: DocumentNode): string {
    return (node.content ?? [])
      .map((child) => this.inline(child.content ?? [], 'cell').trim())
      .filter((text) => text !== '')
      .join('<br>')
  }

  /** O sumário: uma lista de links para os títulos, aninhada pelo nível. */
  private contents(node: DocumentNode): string {
    const children = node.content ?? []
    const head = Math.max(0, Number(node.attrs?.['head']) || 0)
    const title = this.blocks(children.slice(0, head))
    const entries = children
      .slice(head)
      .map((entry) => {
        const text = this.inline(withoutPageNumbers(entry.content ?? []), 'heading').trim()
        return text === '' ? '' : `${'  '.repeat(tocLevelOf(entry) - 1)}- ${text}`
      })
      .filter((line) => line !== '')
      .join('\n')
    return [title, entries].filter((part) => part !== '').join('\n\n')
  }

  inline(nodes: readonly DocumentNode[], mode: Mode): string {
    let out = ''
    const stack: OpenMark[] = []

    const close = (until: number): void => {
      while (stack.length > until) {
        const entry = stack.pop()!
        const trailing = /[ \t]*$/.exec(out)![0]
        const core = out.slice(0, out.length - trailing.length)
        // A marca que ficou sem conteúdo sai sem deixar `****` para trás.
        out =
          core.length === entry.at
            ? core.slice(0, entry.at - entry.open.length) + trailing
            : core + entry.close + trailing
      }
    }

    for (const node of nodes) {
      let piece = this.inlineContent(node, mode)
      if (piece === '') continue
      const isCode = node.type === 'text' && (node.marks ?? []).some((mark) => mark.type === 'code')
      if (isCode) piece = codeSpan(node.text ?? '', mode)

      let wanted = this.marksOf(node)
      // O espaço sozinho não abre ênfase: `** **` não é negrito em Markdown.
      if (node.type === 'text' && !isCode && (node.text ?? '').trim() === '') {
        wanted = wanted.filter((mark) => stack.some((open) => open.key === mark.key))
      }

      const keep = stack.findIndex((open, index) => wanted[index]?.key !== open.key)
      close(keep < 0 ? stack.length : keep)

      const opening = wanted.slice(stack.length)
      if (opening.length > 0) {
        // O espaço antes do conteúdo vai para fora da marca.
        const leading = /^[ \t]*/.exec(piece)![0]
        out += leading
        piece = piece.slice(leading.length)
        for (const mark of opening) {
          out += mark.open
          stack.push({ ...mark, at: out.length })
        }
      }
      out += piece
    }
    close(0)
    return out
  }

  private marksOf(node: DocumentNode): Array<Omit<OpenMark, 'at'>> {
    // A referência de nota fica fora das marcas: `<sup>[^1]</sup>` não é nota.
    if (node.type === 'noteRef') return []
    const found: Array<Omit<OpenMark, 'at'> & { readonly order: number }> = []
    for (const mark of node.marks ?? []) {
      const order = MARK_ORDER.indexOf(mark.type)
      if (order < 0 || mark.attrs?.['off'] === true) continue
      switch (mark.type) {
        case 'link': {
          const href = safeHref(mark)
          if (href === null) continue
          const title = typeof mark.attrs?.['title'] === 'string' ? mark.attrs['title'] : ''
          found.push({ order, key: `link ${href}`, open: '[', close: `](${linkDestination(href, title)})` })
          break
        }
        case 'bold':
          found.push({ order, key: 'bold', open: '**', close: '**' })
          break
        case 'italic':
          found.push({ order, key: 'italic', open: '*', close: '*' })
          break
        case 'strike':
          found.push({ order, key: 'strike', open: '~~', close: '~~' })
          break
        case 'superscript':
          found.push({ order, key: 'sup', open: '<sup>', close: '</sup>' })
          break
        case 'subscript':
          found.push({ order, key: 'sub', open: '<sub>', close: '</sub>' })
          break
      }
    }
    return found.sort((a, b) => a.order - b.order).map(({ key, open, close }) => ({ key, open, close }))
  }

  private inlineContent(node: DocumentNode, mode: Mode): string {
    switch (node.type) {
      case 'text':
        return escapeMarkdown((node.text ?? '').replace(/\r?\n/g, ' '))
      case 'hardBreak':
        return mode === 'block' ? '\\\n' : mode === 'cell' ? '<br>' : ' '
      case 'image': {
        const alt = typeof node.attrs?.['alt'] === 'string' ? node.attrs['alt'] : ''
        const path = this.imagePath(node.attrs?.['src'])
        if (path === null) return escapeMarkdown(alt)
        return `![${escapeMarkdown(alt)}](${path})`
      }
      case 'noteRef': {
        const note = this.source.noteOf.get(node)
        return note === undefined ? '' : `[^${noteKey(note)}]`
      }
      case 'field':
        return escapeMarkdown(String(node.attrs?.['result'] ?? ''))
      // A equação vai em LaTeX entre cifrões — `$…$` no texto, `$$…$$` na linha
      // dela quando é de exibição —, que é o que o Pandoc, o GitHub e os
      // editores de Markdown leem. Sem LaTeX nenhum, vai o MathML como HTML.
      case 'math':
        return mathMarkdown(node, mode)
      case 'bookmarkStart': {
        // Só os que algum link aponta: os outros seriam ruído no texto.
        const name = String(node.attrs?.['name'] ?? '')
        return this.source.linkTargets.has(name) ? `<a id="${escapeHtml(name)}"></a>` : ''
      }
      default:
        return node.content === undefined ? '' : this.inline(node.content, mode)
    }
  }
}

function mathMarkdown(node: DocumentNode, mode: Mode): string {
  const latex = latexOfEquation(node.attrs).replace(/\s*\n\s*/g, ' ')
  if (latex === '') return mathHtml(node)
  const escaped = mode === 'cell' ? latex.replace(/\|/g, '\\|') : latex
  if (node.attrs?.['display'] !== true) return `$${escaped}$`
  return mode === 'block' ? `\n$$${escaped}$$\n` : `$$${escaped}$$`
}

/** `[^1]` nas de rodapé, `[^fim-1]` nas de fim: as duas contas não se cruzam. */
function noteKey(note: ExportNote): string {
  return note.id.replace(/^nota-rodape-/, '').replace(/^nota-/, '')
}

function isList(node: DocumentNode): boolean {
  return node.type === 'bulletList' || node.type === 'orderedList'
}

/** A tabela que o GFM representa: sem mescla, sem bloco dentro de célula. */
function isSimpleTable(rows: readonly DocumentNode[]): boolean {
  const columns = (rows[0]?.content ?? []).length
  if (columns === 0) return false
  return rows.every((row) => {
    const cells = row.content ?? []
    if (cells.length !== columns) return false
    return cells.every((cell) => {
      const span = (name: string): number => Number(cell.attrs?.[name] ?? 1) || 1
      if (span('colspan') > 1 || span('rowspan') > 1) return false
      return (cell.content ?? []).every((child) => child.type === 'paragraph' || child.type === 'heading')
    })
  })
}

/**
 * Escapa o que o Markdown leria como marcação no meio da linha. O começo da
 * linha (`#`, `>`, `-`, `1.`) é com `escapeLineStarts`, que sabe onde ela começa.
 */
export function escapeMarkdown(text: string): string {
  return text.replace(/[\\`*_[\]<>~|$]/g, '\\$&').replace(/&(?=#?[a-z0-9]+;)/gi, '&amp;')
}

/** Escapa, em cada linha do parágrafo, o que no começo dela abriria outro bloco. */
function escapeLineStarts(text: string): string {
  return text
    .split('\n')
    .map((line) => {
      // Recuo no começo viraria bloco de código; no Markdown ele não significa nada.
      const trimmed = line.replace(/^[ \t]+/, '')
      if (/^(#{1,6}(\s|$)|>|[-+](\s|$)|=+\s*$|-+\s*$)/.test(trimmed)) return `\\${trimmed}`
      return trimmed.replace(/^(\d{1,9})([.)])(?=\s|$)/, '$1\\$2')
    })
    .join('\n')
}

function codeSpan(text: string, mode: Mode): string {
  const content = text.replace(/\r?\n/g, ' ')
  const longest = Math.max(0, ...Array.from(content.matchAll(/`+/g), (match) => match[0].length))
  const fence = '`'.repeat(longest + 1)
  const padded = content.startsWith('`') || content.endsWith('`') ? ` ${content} ` : content
  // Na célula do GFM a barra vertical divide a coluna mesmo dentro do código.
  return `${fence}${mode === 'cell' ? padded.replace(/\|/g, '\\|') : padded}${fence}`
}

function codeBlock(node: DocumentNode): string {
  const text = plainText(node)
  const longest = Math.max(2, ...Array.from(text.matchAll(/`+/g), (match) => match[0].length))
  const fence = '`'.repeat(longest + 1)
  const language =
    typeof node.attrs?.['language'] === 'string' ? node.attrs['language'].replace(/[^\w+-]/g, '') : ''
  return `${fence}${language}\n${text}\n${fence}`
}

function plainText(node: DocumentNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (node.type === 'hardBreak') return '\n'
  return (node.content ?? []).map(plainText).join('')
}

function linkDestination(href: string, title: string): string {
  const destination = /[\s()<>]/.test(href) ? `<${href.replace(/[<>]/g, '\\$&')}>` : href
  return title === '' ? destination : `${destination} "${title.replace(/["\\]/g, '\\$&')}"`
}

/** Um pedaço de caminho de URL relativa: espaço, acento e parêntese codificados. */
function encodePathPart(part: string): string {
  return encodeURIComponent(part).replace(/\(/g, '%28').replace(/\)/g, '%29')
}

/** Recua as linhas depois da primeira — a continuação de um item ou de uma nota. */
function indentRest(text: string, pad: string): string {
  return text
    .split('\n')
    .map((line, index) => (index === 0 || line === '' ? line : pad + line))
    .join('\n')
}

function prefixLines(text: string, prefix: string, blank: string): string {
  return text
    .split('\n')
    .map((line) => (line === '' ? blank : prefix + line))
    .join('\n')
}
