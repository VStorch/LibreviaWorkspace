/**
 * The same CSS goes to the editor and to the HTML that renders the PDF, so paper matches the
 * screen. No app CSS variables: it must work in a standalone HTML.
 */
import { DOCUMENT_FONT_CSS } from './fonts.js'
import { DECLARED_BLACK_VARIABLE } from './style-css.js'

export const DOCUMENT_CONTENT_CSS = `
${DOCUMENT_FONT_CSS}

.page__content {
  outline: none;
  /* Flex, not normal flow, so the margins do not collapse: CSS keeps the larger
     one, while Word and LibreOffice add both. */
  display: flex;
  flex-direction: column;
  /* In the document, the style CSS (style-css.ts) overrides this; it applies on
     its own in the printed spreadsheet. */
  font-family: 'Times New Roman', 'Liberation Serif', Georgia, serif;
  font-size: 12pt;
  line-height: 1.5;
  color: #111111;
}

.page__content > * + * { margin-top: 0.6em; }

/* An empty paragraph takes one line: on paper there is no <br> that ProseMirror
   adds for the cursor, and without this the text would move up. */
.page__content :is(p, h1, h2, h3, h4, h5, h6, li):empty::before {
  content: '';
  display: inline-block;
}

.page__content h1 { font-size: 22pt; font-weight: 600; margin-top: 1em; }
.page__content h2 { font-size: 17pt; font-weight: 600; margin-top: 1em; }
.page__content h3 { font-size: 14pt; font-weight: 600; margin-top: 1em; }
.page__content h4 { font-size: 12pt; font-weight: 700; margin-top: 1em; }

.page__content ul,
.page__content ol { padding-left: 1.6em; }

.page__content li > p { margin: 0; }

.page__content a { color: #14538f; text-decoration: underline; }

/* The table of contents: Word's "toc N" style asks for a right tab with dot
   leaders. With no tab stops in the editor, the entry becomes a flex row. The
   entry's link is not colored, as in Word. */
.page__content .toc a { color: inherit; text-decoration: none; }
.page__content .toc :is(p, h1, h2, h3, h4, h5, h6, a):has(> .field[data-field='pageref']) {
  display: flex;
  align-items: baseline;
}
.page__content .toc .field[data-field='pageref'] {
  flex: 1 1 auto;
  display: flex;
  min-width: 1.5em;
  white-space: nowrap;
}
.page__content .toc .field[data-field='pageref']::before {
  content: '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .';
  flex: 1 1 auto;
  overflow: hidden;
  padding: 0 0.2em;
}

/* The math font is the system's: \`math\` is the generic family Chromium
   resolves to whatever font has an OpenType MATH table. The box
   (\`m:borderBox\`) is an \`mrow\` with a border, because MathML Core has no
   \`menclose\`. */
.page__content .equation math {
  font-family: 'STIX Two Math', 'Cambria Math', 'Noto Sans Math', 'DejaVu Math TeX Gyre', math;
}
.page__content .equation--display { display: block; text-align: center; break-inside: avoid; }
.page__content .equation--display[data-jc='left'] { text-align: left; }
.page__content .equation--display[data-jc='right'] { text-align: right; }
.page__content .equation--display math { display: inline math; math-style: normal; }
.page__content .equation .omml-caixa { border: 1px solid currentColor; padding: 0.1em; }

/* Superscript and subscript without stretching the line, as in Word:
   \`line-height: 0\` keeps the line's height, and the fractional size does not
   shrink again with each nested exponent. */
.page__content sup,
.page__content sub { font-size: 0.65em; line-height: 0; }

.page__content img { max-width: 100%; height: auto; }

/* Image as a block, aligned by attribute: only comes from old drafts. */
.page__content img[data-align='center'] { display: block; margin-inline: auto; }
.page__content img[data-align='right'] { display: block; margin-left: auto; }

/* The list mark comes ready in --list-marker (list-numbering.ts), counted the
   way Word counts, and is drawn by a pseudo-element because the native marker
   cannot be placed in the hanging indent (w:ind/@hanging). The indent is
   relative to the outer list's (--list-indent); !important because the node may
   carry the absolute one as an inline style. */
.page__content ul[data-list-indent],
.page__content ol[data-list-indent] {
  padding-left: var(--list-indent) !important;
  margin-left: var(--list-margin, 0mm);
}

.page__content li[data-label] { list-style: none; }

.page__content li[data-label] > :first-child::before {
  content: var(--list-marker);
  display: inline-block;
  box-sizing: border-box;
  min-width: var(--list-hanging, 0mm);
  margin-left: calc(-1 * var(--list-hanging, 0mm));
  /* A mark wider than the hanging indent: Word jumps to the next tab stop;
     here, at least half a letter of space. */
  padding-right: 0.5em;
  text-indent: 0;
  white-space: pre;
}

/* An anchored image is a column-wide block, as in Word: inline it would pay
   for the font's descent, and the paragraph indent does not narrow it. By
   attribute, not by position, because ProseMirror puts a service image next to
   it. On screen the rule targets the NodeView wrapper (.node-image), a child of
   the paragraph; the frame shrinks to the image so the handles sit on the
   corners. */
.page__content img[data-anchored],
.page__content .node-image[data-anchored] {
  display: block;
  margin-left: calc(-1 * var(--indent, 0mm));
  max-width: calc(100% + var(--indent, 0mm) + var(--indent-right, 0mm));
}

.page__content .node-image[data-anchored] > .image-frame {
  display: block;
  width: fit-content;
}

/* ProseMirror's service <br> does not open a line after a block image: paper
   has no such <br>. */
.page__content img[data-anchored] ~ br,
.page__content img[data-anchored] ~ img.ProseMirror-separator,
.page__content .node-image[data-anchored] ~ br,
.page__content .node-image[data-anchored] ~ img.ProseMirror-separator { display: none !important; }

/* The anchored image's paragraph keeps its line under the image, inside the
   paragraph: that is what LibreOffice measures between two captures. */
.page__content p:not([data-anchor-text]):has(> img[data-anchored])::after,
.page__content p:not([data-anchor-text]):has(> .node-image[data-anchored])::after {
  content: '';
  display: block;
  height: 1lh;
}


/* The section mark (the empty paragraph holding w:sectPr) takes no line, as in
   LibreOffice. It stays in the document: saving writes it back. */
.page__content p[data-section-mark] {
  height: 0;
  margin: 0;
  overflow: hidden;
}

.page__content p[data-section-mark]::before { content: none; }

.page__content blockquote {
  border-left: 3px solid #c7ced6;
  padding-left: 1em;
  color: #444444;
}

.page__content table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
  /* No margin of its own, as in Word: the space around it is the neighboring
   paragraphs'. */
  margin: 0;
}

/* The cell margin comes from the table (--cell-margins); without it, 60/120
   twips, Word's TableNormal. */
.page__content th,
.page__content td {
  border: 1px solid #9aa3ad;
  padding: var(--cell-margins, 4px 8px);
  vertical-align: top;
  position: relative;
}

.page__content th {
  background: #f0f2f4;
  font-weight: 600;
  text-align: left;
}
`

/** The dark theme lives here, never in `DOCUMENT_CONTENT_CSS`: printed paper is white in any theme. */
const DARK_CONTENT_CSS = `
:root[data-theme='dark'] .page__content {
  color: var(--text);
  ${DECLARED_BLACK_VARIABLE}: var(--text);
}

/* The same color as the other surfaces, so the paper does not look cut out. */
:root[data-theme='dark'] .paper {
  background: var(--surface);
}

/* #14538f on dark paper gives 2.1:1. */
:root[data-theme='dark'] .page__content a {
  color: var(--accent-document);
}

:root[data-theme='dark'] .page__content blockquote {
  border-left-color: var(--border-strong);
  color: var(--muted);
}

:root[data-theme='dark'] .page__content th,
:root[data-theme='dark'] .page__content td {
  border-color: var(--border-strong);
}

:root[data-theme='dark'] .page__content th {
  background: var(--hover);
}

:root[data-theme='dark'] .page__content .tiptap-invisible-character::before {
  color: var(--muted);
}
`

/* The separator height comes from pagination; the line sits in its middle. */
export const NOTES_CSS = `
.paper-notes { position: absolute; }
.paper-notes__separator { position: relative; }
.paper-notes__separator::after {
  content: '';
  position: absolute;
  left: 0;
  top: 50%;
  width: 33%;
  border-top: 0.75pt solid #000000;
}
.paper-notes__separator--continued::after { width: 100%; }
/* clip, not hidden: the area does not scroll with the cursor. flow-root because
   clip does not contain the continuation's negative margin. */
.paper-notes__slot { overflow: clip; display: flow-root; }
.note-body { display: flex; flex-direction: column; }
.note-number { vertical-align: super; font-size: 0.65em; line-height: 0; }
`

/** Editor only: cell selection, resize handle, page break. */
export const EDITOR_ONLY_CSS = `
/* Word's gray shading shows that the number is calculated. Screen only. */
.page__content .field.ProseMirror-selectednode { background: #d9d9d9; outline: none; }

/* The number is decoration, from the order in the text (note-ref.ts); on paper
   it is written out. */
.page__content .note-ref::after { content: attr(data-note-number); }
.page__content .note-ref.ProseMirror-selectednode { background: #d9d9d9; outline: none; }

/* A locked equation, with a construct the screen cannot draw, gets an
   underline, and the tooltip says what is missing. */
.page__content .equation.ProseMirror-selectednode { background: #d9d9d9; outline: none; }
.page__content .equation--locked { text-decoration: underline dotted #b0b0b0; }

${DARK_CONTENT_CSS}
.page__content .selectedCell::after {
  content: '';
  position: absolute;
  inset: 0;
  background: rgb(31 95 169 / 12%);
  pointer-events: none;
}

.page__content .column-resize-handle {
  position: absolute;
  right: -2px;
  top: 0;
  bottom: 0;
  width: 4px;
  background: #1f5fa9;
  cursor: col-resize;
}

/* A page break draws nothing, as in the print layout of Word and LibreOffice:
   the sheet ends there. It stays selectable and deletable with Backspace at the
   start of the next sheet. */
.page__content .page-break {
  border: none;
  margin: 0;
  height: 0;
}

.page__content .ProseMirror-selectednode { outline: 2px solid #1f5fa9; }

/* Inline and with no line spacing: without this the frame would pay for the
   font's descent, and the screen would measure the image taller than paper. */
.page__content .image-frame {
  display: inline-block;
  position: relative;
  max-width: 100%;
  line-height: 0;
}

.page__content .image-frame > img { max-width: 100%; }

.page__content .image-frame--resizing > img { opacity: 0.75; }

.page__content .image-frame__grip {
  position: absolute;
  width: 10px;
  height: 10px;
  padding: 0;
  border: 1px solid #ffffff;
  background: #1f5fa9;
  z-index: 2;
}

.page__content .image-frame__grip--nw { left: -5px; top: -5px; }
.page__content .image-frame__grip--n { left: calc(50% - 5px); top: -5px; }
.page__content .image-frame__grip--ne { right: -5px; top: -5px; }
.page__content .image-frame__grip--e { right: -5px; top: calc(50% - 5px); }
.page__content .image-frame__grip--se { right: -5px; bottom: -5px; }
.page__content .image-frame__grip--s { left: calc(50% - 5px); bottom: -5px; }
.page__content .image-frame__grip--sw { left: -5px; bottom: -5px; }
.page__content .image-frame__grip--w { left: -5px; top: calc(50% - 5px); }

/* Formatting marks do not change the line's measure: zero width, height and
   line spacing, with the glyph visible through overflow. Hence injectCSS: false
   in editor-extensions.ts. */
.page__content .tiptap-invisible-character {
  width: 0;
  height: 0;
  padding: 0;
  line-height: 0;
  pointer-events: none;
  user-select: none;
}

.page__content .tiptap-invisible-character::before {
  display: inline-block;
  width: 0;
  line-height: 0;
  color: #9aa3ad;
  font-style: normal;
  font-weight: 400;
  caret-color: inherit;
}

.page__content .tiptap-invisible-character--space::before { content: '·'; }
.page__content .tiptap-invisible-character--tab::before { content: '→'; }
.page__content .tiptap-invisible-character--break::before { content: '¬'; }
.page__content .tiptap-invisible-character--paragraph::before { content: '¶'; }

/* The service image next to the mark takes no height either. */
.page__content .tiptap-invisible-character + img.ProseMirror-separator {
  width: 0 !important;
  height: 0 !important;
  pointer-events: none;
  user-select: none;
}
`

/** Paper only: orphan heading, split table row and cropped image. */
export const PRINT_ONLY_CSS = `
.page__content [data-page-break] {
  break-after: page;
  border: none;
  height: 0;
  margin: 0;
}

.page__content h1,
.page__content h2,
.page__content h3,
.page__content h4,
/* The same rule as the screen's end-of-page marks. */
.page__content [data-keep-next] { break-after: avoid; }

.page__content tr,
.page__content img { break-inside: avoid; }

.page__content thead { display: table-header-group; }

/* As on screen (styles.css); in the other modes printing goes without the marks
   (print-source.ts). */
.page__content .revision-author-0 { --revision: #1f5fa9; }
.page__content .revision-author-1 { --revision: #b3261e; }
.page__content .revision-author-2 { --revision: #1a7a4c; }
.page__content .revision-author-3 { --revision: #8a4baf; }
.page__content .revision-author-4 { --revision: #b26a00; }
.page__content .revision-author-5 { --revision: #00796b; }
.page__content ins.revision {
  color: var(--revision);
  text-decoration: underline;
  text-decoration-color: var(--revision);
}
.page__content del.revision {
  color: var(--revision);
  text-decoration: line-through;
  text-decoration-color: var(--revision);
}
.page__content [data-revision]::after { content: '¶'; color: var(--revision); line-height: 0; }
.page__content [data-revision='del']::after { text-decoration: line-through; }
.page__content tr[data-revision] { box-shadow: inset 3px 0 0 var(--revision); }
.page__content tr[data-revision]::after { content: none; }
.page__content tr[data-revision='del'] :is(td, th) { text-decoration: line-through; }
`
