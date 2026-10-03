import {
  exportTitle,
  imageData,
  isSectionMarkOnly,
  prepareExport,
  safeHref,
  tocLevelOf,
  withoutPageNumbers,
  type ExportSource,
} from './export-common.js'
import { itemDrawAttrs, listDrawAttrs } from './list-numbering.js'
import { mathMlToString, sanitizeMathMl } from './mathml.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { NoteKind } from './notes.js'
import { styleSheetCss } from './style-css.js'
import { cellBordersFromAttr, cellBordersToCss } from './table-format.js'

/**
 * Uma página autocontida e sem script. Os estilos viram as regras da tela
 * (`styleSheetCss`) e a formatação direta sai inline. Tudo o que vem do
 * documento é escapado, e o CSS de atributo passa por `safeCss`.
 */

export interface HtmlExportOptions {
  /** O nome do arquivo de origem: o título quando as propriedades não têm um. */
  readonly fileName: string
  readonly lang: string
  /** Os textos que a página leva, na língua da interface. */
  readonly labels: HtmlExportLabels
}

export interface HtmlExportLabels {
  /** O rótulo da seção de notas, para leitores de tela. */
  readonly notes: string
  /** A dica do número da nota, que leva de volta ao texto. */
  readonly backToText: string
}

/** Como o renderizador resolve o `src` de uma imagem: `null` a omite. */
export type ImageSource = (src: unknown) => string | null

/** O Markdown o usa para a tabela que não cabe em GFM. */
export interface HtmlRenderer {
  blocks(nodes: readonly DocumentNode[]): string
  inline(nodes: readonly DocumentNode[]): string
}

const EMBEDDED_IMAGE: ImageSource = (src) => (imageData(src) === null ? null : String(src))

export function exportHtml(
  model: Pick<DocumentModel, 'doc' | 'styles' | 'notes' | 'properties'>,
  options: HtmlExportOptions,
): string {
  const source = prepareExport(model)
  const renderer = createHtmlRenderer(source, EMBEDDED_IMAGE)
  const body = renderer.blocks(source.doc.content ?? [])
  const notes = notesSection(source, renderer, options.labels)
  const properties = model.properties

  const meta = [
    metaTag('author', properties?.creator),
    metaTag('description', properties?.subject || properties?.description),
    metaTag('keywords', properties?.keywords),
  ].join('')

  return `<!doctype html>
<html lang="${escapeHtml(options.lang)}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">
<meta name="generator" content="Librevia">
${meta}<title>${escapeHtml(exportTitle(model, options.fileName))}</title>
<style>
${EXPORT_CSS}
${safeStyleText(styleSheetCss(model.styles))}
</style>
</head>
<body>
<main class="page__content">
${body}
</main>
${notes}</body>
</html>
`
}

function metaTag(name: string, value: string | undefined): string {
  const text = value?.trim()
  return text === undefined || text === '' ? '' : `<meta name="${name}" content="${escapeHtml(text)}">\n`
}

/** A página não tem a folha do editor: uma coluna legível, com as regras de lista e tabela dele. */
const EXPORT_CSS = `
html { background: #ffffff; }
body { margin: 0; color: #111111; }
.page__content {
  display: flex;
  flex-direction: column;
  box-sizing: border-box;
  max-width: 170mm;
  margin: 24px auto;
  padding: 0 16px;
}
.page__content :is(p, h1, h2, h3, h4, h5, h6) { white-space: pre-wrap; overflow-wrap: break-word; }
.page__content li > p { margin: 0; }
.page__content a { color: #14538f; }
.page__content ul[data-list-indent],
.page__content ol[data-list-indent] {
  padding-left: var(--lista-recuo) !important;
  margin-left: var(--lista-margem, 0mm);
}
.page__content li[data-label] { list-style: none; }
.page__content li[data-label] > :first-child::before {
  content: var(--lista-marca);
  display: inline-block;
  box-sizing: border-box;
  min-width: var(--lista-pendente, 0mm);
  margin-left: calc(-1 * var(--lista-pendente, 0mm));
  padding-right: 0.5em;
  text-indent: 0;
  white-space: pre;
}
.page__content table { border-collapse: collapse; width: 100%; margin: 0; }
.page__content th,
.page__content td { border: 1px solid #9aa3ad; padding: 4px 8px; vertical-align: top; }
.page__content th { background: #f0f2f4; font-weight: 600; text-align: left; }
.page__content img { max-width: 100%; height: auto; }
.page__content blockquote { border-left: 3px solid #c7ced6; padding-left: 1em; color: #444444; }
.sumario ul { list-style: none; margin: 0; padding: 0; }
.sumario a { color: inherit; text-decoration: none; }
.nota-ref a { text-decoration: none; }
.notas { box-sizing: border-box; max-width: 170mm; margin: 0 auto 24px; padding: 0 16px; font-size: 0.9em; }
.notas hr { width: 33%; margin: 0 0 0.6em; border: 0; border-top: 1px solid #555555; }
.notas ol { list-style: none; margin: 0; padding: 0; }
.notas li { display: flex; gap: 0.4em; }
.notas li > a { text-decoration: none; }
.notas__corpo > * { margin: 0; }
.page-break { display: none; }
@media print {
  .page__content, .notas { max-width: none; margin: 0; padding: 0; }
  .page-break { display: block; break-after: page; }
}
`

export function createHtmlRenderer(source: ExportSource, imageSrc: ImageSource): HtmlRenderer {
  function blocks(nodes: readonly DocumentNode[]): string {
    return nodes
      .map(block)
      .filter((html) => html !== '')
      .join('\n')
  }

  function block(node: DocumentNode): string {
    switch (node.type) {
      case 'paragraph': {
        if (isSectionMarkOnly(node)) return ''
        return `<p${blockAttrs(node)}>${inlineOrBreak(node.content)}</p>`
      }
      case 'heading': {
        const level = Math.min(6, Math.max(1, Number(node.attrs?.['level']) || 1))
        return `<h${level}${blockAttrs(node)}>${inlineOrBreak(node.content)}</h${level}>`
      }
      case 'bulletList':
      case 'orderedList':
        return list(node)
      case 'table':
        return table(node)
      case 'blockquote':
        return `<blockquote>\n${blocks(node.content ?? [])}\n</blockquote>`
      case 'codeBlock':
        return `<pre><code>${escapeHtml(plainText(node))}</code></pre>`
      case 'horizontalRule':
        return '<hr>'
      case 'pageBreak':
        return '<div class="page-break"></div>'
      case 'tableOfContents':
        return contents(node)
      default:
        return node.content === undefined ? '' : blocks(node.content)
    }
  }

  function inlineOrBreak(content: readonly DocumentNode[] | undefined): string {
    const html = inline(content ?? [])
    // O parágrafo vazio ocupa uma linha, como no documento.
    return html === '' ? '<br>' : html
  }

  function blockAttrs(node: DocumentNode): string {
    const attrs = node.attrs ?? {}
    const style = attrs['styleId']
    const parts: string[] = []
    if (typeof style === 'string' && style !== '') parts.push(` data-style-id="${escapeHtml(style)}"`)
    const css = blockCss(attrs)
    if (css !== '') parts.push(` style="${escapeHtml(css)}"`)
    return parts.join('')
  }

  function list(node: DocumentNode): string {
    const info = source.listOf.get(node)
    const ordered = node.type === 'orderedList'
    const tag = ordered ? 'ol' : 'ul'
    const items = (node.content ?? []).filter((child) => child.type === 'listItem')
    const first = items[0] === undefined ? undefined : source.itemOf.get(items[0])
    const attrs: string[] = []
    if (ordered && first !== undefined && first.value !== 1) attrs.push(` start="${first.value}"`)
    if (info !== undefined) {
      for (const [name, value] of Object.entries(listDrawAttrs(info))) {
        attrs.push(value === '' ? ` ${name}` : ` ${name}="${escapeHtml(value)}"`)
      }
    }
    const body = items
      .map((item) => {
        const label = source.itemOf.get(item)?.label
        const drawn =
          label === undefined
            ? ''
            : Object.entries(itemDrawAttrs(label))
                .map(([name, value]) => ` ${name}="${escapeHtml(value)}"`)
                .join('')
        return `<li${drawn}>${blocks(item.content ?? [])}</li>`
      })
      .join('\n')
    return `<${tag}${attrs.join('')}>\n${body}\n</${tag}>`
  }

  function table(node: DocumentNode): string {
    const rows = (node.content ?? []).filter((row) => row.type === 'tableRow')
    const widths = columnWidths(rows[0])
    const colgroup =
      widths === null
        ? ''
        : `<colgroup>${widths.map((width) => `<col style="width: ${width}px">`).join('')}</colgroup>`
    const body = rows.map((row) => `<tr>${(row.content ?? []).map(cell).join('')}</tr>`).join('\n')
    return `<table>${colgroup}\n${body}\n</table>`
  }

  function cell(node: DocumentNode): string {
    const tag = node.type === 'tableHeader' ? 'th' : 'td'
    const attrs = node.attrs ?? {}
    const parts: string[] = []
    const colspan = positive(attrs['colspan'])
    const rowspan = positive(attrs['rowspan'])
    if (colspan !== null && colspan > 1) parts.push(` colspan="${colspan}"`)
    if (rowspan !== null && rowspan > 1) parts.push(` rowspan="${rowspan}"`)
    const css: string[] = []
    css.push(...safeDeclarations(cellBordersToCss(cellBordersFromAttr(attrs['borders']))))
    const shading = safeCss(attrs['shading'])
    if (shading !== null) css.push(`background-color: ${shading}`)
    if (css.length > 0) parts.push(` style="${escapeHtml(css.join('; '))}"`)
    return `<${tag}${parts.join('')}>${blocks(node.content ?? [])}</${tag}>`
  }

  function contents(node: DocumentNode): string {
    const children = node.content ?? []
    const head = Math.max(0, Number(node.attrs?.['head']) || 0)
    const title = blocks(children.slice(0, head))
    const entries = children
      .slice(head)
      .map((entry) => {
        const level = tocLevelOf(entry)
        const indent = level > 1 ? ` style="margin-left: ${(level - 1) * 1.5}em"` : ''
        return `<li${indent}>${inline(withoutPageNumbers(entry.content ?? []))}</li>`
      })
      .join('\n')
    const nav = `<nav class="sumario"><ul>\n${entries}\n</ul></nav>`
    return title === '' ? nav : `${title}\n${nav}`
  }

  function inline(nodes: readonly DocumentNode[]): string {
    return nodes.map(inlineNode).join('')
  }

  function inlineNode(node: DocumentNode): string {
    const inner = inlineContent(node)
    if (inner === '') return ''
    // O `vertAlign` que o Word põe na referência de nota a subiria duas vezes.
    return node.type === 'noteRef' ? inner : wrapMarks(inner, node.marks ?? [])
  }

  function inlineContent(node: DocumentNode): string {
    switch (node.type) {
      case 'text':
        return escapeHtml(node.text ?? '')
      case 'hardBreak':
        return '<br>'
      case 'image': {
        const src = imageSrc(node.attrs?.['src'])
        const alt = stringAttr(node.attrs?.['alt'])
        if (src === null) return alt === '' ? '' : escapeHtml(alt)
        const parts = [`<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}"`]
        const title = stringAttr(node.attrs?.['title'])
        if (title !== '') parts.push(` title="${escapeHtml(title)}"`)
        const width = positive(node.attrs?.['width'])
        const height = positive(node.attrs?.['height'])
        if (width !== null) parts.push(` width="${width}"`)
        if (height !== null) parts.push(` height="${height}"`)
        if (node.attrs?.['anchored'] === true) parts.push(' style="display: block"')
        return `${parts.join('')}>`
      }
      case 'noteRef': {
        const note = source.noteOf.get(node)
        if (note === undefined) return ''
        const kind = note.kind === NoteKind.Endnote ? ' data-kind="endnote"' : ''
        return `<sup class="nota-ref"${kind}><a href="#${note.id}" id="ref-${note.id}">${escapeHtml(note.label)}</a></sup>`
      }
      case 'field':
        return escapeHtml(stringAttr(node.attrs?.['result']))
      case 'math':
        return mathHtml(node)
      case 'bookmarkStart': {
        const name = stringAttr(node.attrs?.['name'])
        return name === '' ? '' : `<a id="${escapeHtml(name)}"></a>`
      }
      default:
        return node.content === undefined ? escapeHtml(node.text ?? '') : inline(node.content)
    }
  }

  return { blocks, inline }
}

function notesSection(source: ExportSource, renderer: HtmlRenderer, labels: HtmlExportLabels): string {
  if (source.notes.length === 0) return ''
  const lists = [NoteKind.Footnote, NoteKind.Endnote]
    .map((kind) => source.notes.filter((note) => note.kind === kind))
    .filter((notes) => notes.length > 0)
    .map(
      (notes) =>
        `<ol>\n${notes
          .map(
            (note) =>
              `<li id="${note.id}"><a href="#ref-${note.id}" title="${escapeHtml(labels.backToText)}">${escapeHtml(note.label)}</a>` +
              `<div class="notas__corpo">${renderer.blocks(note.body)}</div></li>`,
          )
          .join('\n')}\n</ol>`,
    )
  return `<section class="notas" aria-label="${escapeHtml(labels.notes)}">\n<hr>\n${lists.join('\n')}\n</section>\n`
}

/** As marcas de um trecho, de fora para dentro: o link por fora de tudo. */
function wrapMarks(inner: string, marks: NonNullable<DocumentNode['marks']>): string {
  let html = inner
  const ordered = [...marks].sort((a, b) => (a.type === 'link' ? 1 : 0) - (b.type === 'link' ? 1 : 0))
  for (const mark of ordered) html = wrapMark(html, mark)
  return html
}

function wrapMark(html: string, mark: NonNullable<DocumentNode['marks']>[number]): string {
  const attrs = mark.attrs ?? {}
  const off = attrs['off'] === true
  switch (mark.type) {
    case 'bold':
      return off ? `<span style="font-weight: 400">${html}</span>` : `<strong>${html}</strong>`
    case 'italic':
      return off ? `<span style="font-style: normal">${html}</span>` : `<em>${html}</em>`
    case 'underline':
      return off ? `<span style="text-decoration: none">${html}</span>` : `<u>${html}</u>`
    case 'strike':
      return off ? `<span style="text-decoration: none">${html}</span>` : `<s>${html}</s>`
    case 'code':
      return `<code>${html}</code>`
    case 'superscript':
      return `<sup>${html}</sup>`
    case 'subscript':
      return `<sub>${html}</sub>`
    case 'caps':
      return `<span style="text-transform: uppercase">${html}</span>`
    case 'smallCaps':
      return `<span style="font-variant: small-caps">${html}</span>`
    case 'highlight': {
      const color = safeCss(attrs['color'])
      return color === null
        ? `<mark>${html}</mark>`
        : `<mark style="background-color: ${escapeHtml(color)}">${html}</mark>`
    }
    case 'charStyle': {
      const id = attrs['styleId']
      return typeof id === 'string' && id !== ''
        ? `<span data-char-style="${escapeHtml(id)}">${html}</span>`
        : html
    }
    case 'textStyle': {
      const css = textStyleCss(attrs)
      return css === '' ? html : `<span style="${escapeHtml(css)}">${html}</span>`
    }
    case 'link': {
      const href = safeHref(mark)
      if (href === null) return html
      const title =
        typeof attrs['title'] === 'string' && attrs['title'] !== ''
          ? ` title="${escapeHtml(attrs['title'])}"`
          : ''
      const external = href.startsWith('#') ? '' : ' rel="noopener noreferrer"'
      return `<a href="${escapeHtml(href)}"${title}${external}>${html}</a>`
    }
    default:
      return html
  }
}

function textStyleCss(attrs: Record<string, unknown>): string {
  const pairs: Array<[string, unknown]> = [
    ['color', attrs['color']],
    ['background-color', attrs['backgroundColor']],
    ['font-family', attrs['fontFamily']],
    ['font-size', attrs['fontSize']],
    ['line-height', attrs['lineHeight']],
  ]
  return declarations(pairs)
}

/** Como `block-format.ts` a desenha. */
export function blockCss(attrs: Record<string, unknown>): string {
  const mm = (value: unknown): string | null => {
    const number = finiteNumber(value)
    return number === null || number === 0 ? null : `${number}mm`
  }
  const hanging = finiteNumber(attrs['hangingMm'])
  const indentLevel = finiteNumber(attrs['indent'])
  const pairs: Array<[string, unknown]> = [
    ['text-align', attrs['textAlign']],
    ['padding-left', mm(attrs['indentMm'])],
    ['padding-right', mm(attrs['indentRightMm'])],
    ['text-indent', hanging !== null && hanging > 0 ? `${-hanging}mm` : mm(attrs['firstLineMm'])],
    ['margin-left', indentLevel !== null && indentLevel > 0 ? `${indentLevel * INDENT_STEP_EM}em` : null],
    ['margin-top', attrs['spaceBefore']],
    ['margin-bottom', attrs['spaceAfter']],
    ['font-family', attrs['fontFamily']],
    ['font-size', attrs['fontSize']],
    ['line-height', attrs['lineHeight']],
    ['background-color', attrs['background']],
    ['break-after', attrs['breakAfter'] === true ? 'page' : null],
  ]
  return declarations(pairs)
}

/** O passo do recuo por nível, o mesmo do editor (`indent.ts`). */
const INDENT_STEP_EM = 2.5

function declarations(pairs: ReadonlyArray<[string, unknown]>): string {
  return pairs
    .map(([property, value]) => {
      const safe = safeCss(value)
      return safe === null ? null : `${property}: ${safe}`
    })
    .filter((declaration): declaration is string => declaration !== null)
    .join('; ')
}

/** Sem `url(`, `expression`, barra invertida nem ponto e vírgula. */
export function safeCss(value: unknown): string | null {
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : null
  if (typeof value !== 'string') return null
  const text = value.trim()
  if (text === '' || text.length > 200) return null
  if (!/^[\w\s.,%'"#()+-]*$/.test(text)) return null
  if (/url\s*\(|expression|image-set|@import|attr\s*\(/i.test(text)) return null
  return text
}

function safeDeclarations(css: string): string[] {
  return css
    .split(';')
    .map((declaration) => {
      const colon = declaration.indexOf(':')
      if (colon < 0) return null
      const property = declaration.slice(0, colon).trim()
      const value = safeCss(declaration.slice(colon + 1))
      return /^[a-z-]+$/.test(property) && value !== null ? `${property}: ${value}` : null
    })
    .filter((declaration): declaration is string => declaration !== null)
}

/** O `</style>` não pode aparecer no texto de uma regra: fecharia o bloco. */
function safeStyleText(css: string): string {
  return css.replace(/<\//g, '<\\/')
}

function finiteNumber(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function positive(value: unknown): number | null {
  const number = finiteNumber(Array.isArray(value) ? value[0] : value)
  return number !== null && number > 0 ? Math.round(number) : null
}

function stringAttr(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

function plainText(node: DocumentNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (node.type === 'hardBreak') return '\n'
  return (node.content ?? []).map(plainText).join('')
}

function columnWidths(row: DocumentNode | undefined): number[] | null {
  if (row === undefined) return null
  const widths: number[] = []
  for (const cell of row.content ?? []) {
    const declared = cell.attrs?.['colwidth']
    if (!Array.isArray(declared) || declared.length === 0) return null
    for (const width of declared) {
      const number = finiteNumber(width)
      if (number === null || number <= 0) return null
      widths.push(Math.round(number))
    }
  }
  return widths.length > 0 ? widths : null
}

export function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

/** O `display` sai do nó, e não do MathML: é ele que faz da equação de exibição um bloco. */
export function mathHtml(node: DocumentNode): string {
  const tree = sanitizeMathMl(typeof node.attrs?.['mathml'] === 'string' ? node.attrs['mathml'] : '')
  if (tree === null) return ''
  const display = node.attrs?.['display'] === true ? 'block' : 'inline'
  return mathMlToString({ ...tree, attrs: { ...tree.attrs, display } })
}
