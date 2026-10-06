import { escapeHtml } from '@services/html.js'
import {
  exportHeadingLevel,
  imageData,
  imageExtension,
  isSectionMarkOnly,
  plainText,
  prepareExport,
  safeHref,
  tocLevelOf,
  withoutPageNumbers,
  type ExportNote,
  type ExportSource,
  type Mark,
} from './export-common.js'
import { createHtmlRenderer, mathHtml } from './export-html.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { latexOfEquation } from './mathml-latex.js'

/**
 * CommonMark with GFM tables and notes. Images go to a sibling folder (`relatorio_arquivos/`), like
 * Word's "Save as Web Page": as `data:` the text would be unreadable, and many viewers refuse them.
 * A table with merged cells goes out as HTML; underline, color and font are lost.
 */

export interface MarkdownExportOptions {
  /** Next to the file (`relatorio_arquivos`). */
  readonly assetFolder: string
}

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
  /** Where the mark content starts in the output. */
  readonly at: number
}

/** Opening order: the link outside, superscript inside. */
const MARK_ORDER = ['link', 'bold', 'italic', 'strike', 'superscript', 'subscript']

type MarkSyntax = Omit<OpenMark, 'at'>

const SIMPLE_MARK_SYNTAX: Readonly<Record<string, MarkSyntax>> = {
  bold: { key: 'bold', open: '**', close: '**' },
  italic: { key: 'italic', open: '*', close: '*' },
  strike: { key: 'strike', open: '~~', close: '~~' },
  superscript: { key: 'sup', open: '<sup>', close: '</sup>' },
  subscript: { key: 'sub', open: '<sub>', close: '</sub>' },
}

function linkSyntax(mark: Mark): MarkSyntax | null {
  const href = safeHref(mark)
  if (href === null) return null
  const title = typeof mark.attrs?.['title'] === 'string' ? mark.attrs['title'] : ''
  return { key: `link ${href}`, open: '[', close: `](${linkDestination(href, title)})` }
}

export function exportMarkdown(
  model: Pick<DocumentModel, 'doc' | 'notes'>,
  options: MarkdownExportOptions,
): MarkdownExport {
  const source = prepareExport(model)
  const assets: MarkdownAsset[] = []
  const assetBySrc = new Map<string, string>()

  /** Adds the image to the list only once. */
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
      // Two consecutive lists of the same kind would merge into one.
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
      case 'heading':
        return this.heading(node)
      case 'bulletList':
      case 'orderedList':
        return this.list(node)
      case 'table':
        return this.table(node)
      case 'blockquote':
        return this.blockquote(node)
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

  private heading(node: DocumentNode): string {
    const text = this.inline(node.content ?? [], 'heading').trim()
    if (text === '') return ''
    // A trailing `#` would be read as the optional heading closer.
    return `${'#'.repeat(exportHeadingLevel(node))} ${text.replace(/#$/, '\\#')}`
  }

  private blockquote(node: DocumentNode): string {
    const inner = this.blocks(node.content ?? [])
    return inner === '' ? '' : prefixLines(inner, '> ', '>')
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
        // A sublist sticks to the item paragraph: a tight list stays tight.
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
        // A mark left empty goes away without leaving `****` behind.
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
      // A lone space does not open emphasis: `** **` is not bold in Markdown.
      if (node.type === 'text' && !isCode && (node.text ?? '').trim() === '') {
        wanted = wanted.filter((mark) => stack.some((open) => open.key === mark.key))
      }

      const keep = stack.findIndex((open, index) => wanted[index]?.key !== open.key)
      close(keep < 0 ? stack.length : keep)

      const opening = wanted.slice(stack.length)
      if (opening.length > 0) {
        // The space before the content goes outside the mark.
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
    // A note reference stays outside the marks: `<sup>[^1]</sup>` is not a note.
    if (node.type === 'noteRef') return []
    const found: Array<Omit<OpenMark, 'at'> & { readonly order: number }> = []
    for (const mark of node.marks ?? []) {
      const order = MARK_ORDER.indexOf(mark.type)
      if (order < 0 || mark.attrs?.['off'] === true) continue
      const syntax = mark.type === 'link' ? linkSyntax(mark) : SIMPLE_MARK_SYNTAX[mark.type]
      if (syntax !== null && syntax !== undefined) found.push({ order, ...syntax })
    }
    return found.sort((a, b) => a.order - b.order).map(({ key, open, close }) => ({ key, open, close }))
  }

  private inlineContent(node: DocumentNode, mode: Mode): string {
    switch (node.type) {
      case 'text':
        return escapeMarkdown((node.text ?? '').replace(/\r?\n/g, ' '))
      case 'hardBreak':
        return hardBreakMarkdown(mode)
      case 'image':
        return this.imageMarkdown(node)
      case 'noteRef': {
        const note = this.source.noteOf.get(node)
        return note === undefined ? '' : `[^${noteKey(note)}]`
      }
      case 'field':
        return escapeMarkdown(String(node.attrs?.['result'] ?? ''))
      // LaTeX between dollar signs, which Pandoc and GitHub read; without LaTeX, MathML.
      case 'math':
        return mathMarkdown(node, mode)
      case 'bookmarkStart':
        return this.bookmarkAnchor(node)
      default:
        return node.content === undefined ? '' : this.inline(node.content, mode)
    }
  }

  private imageMarkdown(node: DocumentNode): string {
    const alt = typeof node.attrs?.['alt'] === 'string' ? node.attrs['alt'] : ''
    const path = this.imagePath(node.attrs?.['src'])
    if (path === null) return escapeMarkdown(alt)
    return `![${escapeMarkdown(alt)}](${path})`
  }

  /** Only the ones a link points to: the others would be noise in the text. */
  private bookmarkAnchor(node: DocumentNode): string {
    const name = String(node.attrs?.['name'] ?? '')
    return this.source.linkTargets.has(name) ? `<a id="${escapeHtml(name)}"></a>` : ''
  }
}

function hardBreakMarkdown(mode: Mode): string {
  if (mode === 'block') return '\\\n'
  return mode === 'cell' ? '<br>' : ' '
}

function mathMarkdown(node: DocumentNode, mode: Mode): string {
  const latex = latexOfEquation(node.attrs).replace(/\s*\n\s*/g, ' ')
  if (latex === '') return mathHtml(node)
  const escaped = mode === 'cell' ? latex.replace(/\|/g, '\\|') : latex
  if (node.attrs?.['display'] !== true) return `$${escaped}$`
  return mode === 'block' ? `\n$$${escaped}$$\n` : `$$${escaped}$$`
}

/** `[^1]` for footnotes, `[^fim-1]` for endnotes: the two counts do not cross. */
function noteKey(note: ExportNote): string {
  return note.id.replace(/^nota-rodape-/, '').replace(/^nota-/, '')
}

function isList(node: DocumentNode): boolean {
  return node.type === 'bulletList' || node.type === 'orderedList'
}

/** No merged cells and no blocks inside cells. */
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

/** Line starts (`#`, `>`, `-`, `1.`) are handled by `escapeLineStarts`. */
export function escapeMarkdown(text: string): string {
  return text.replace(/[\\`*_[\]<>~|$]/g, '\\$&').replace(/&(?=#?[a-z0-9]+;)/gi, '&amp;')
}

function escapeLineStarts(text: string): string {
  return text
    .split('\n')
    .map((line) => {
      // Leading indentation would become a code block; in Markdown it means nothing.
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
  // In a GFM cell the pipe splits the column even inside code.
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

function linkDestination(href: string, title: string): string {
  const destination = /[\s()<>]/.test(href) ? `<${href.replace(/[<>]/g, '\\$&')}>` : href
  return title === '' ? destination : `${destination} "${title.replace(/["\\]/g, '\\$&')}"`
}

/** Spaces, accents and parentheses encoded. */
function encodePathPart(part: string): string {
  return encodeURIComponent(part).replace(/\(/g, '%28').replace(/\)/g, '%29')
}

/** The continuation of an item or a note. */
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
