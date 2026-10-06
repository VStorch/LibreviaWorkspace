import { escapeHtml } from '@services/html.js'
import { DOCUMENT_CONTENT_CSS, NOTES_CSS, PRINT_ONLY_CSS } from './content-styles.js'

/**
 * The body comes from the editor itself (`editor.getHTML()`). Page size and margins belong to
 * `printToPDF` (`@services/pdf/page-setup.ts`). `extraCss` is for spreadsheets.
 */
export function buildPrintHtml(
  bodyHtml: string,
  title: string,
  extraCss = '',
  /** A paginated document is **not** wrapped: each sheet brings its own `.page__content`. */
  wrapInContent = true,
): string {
  return `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
html, body { margin: 0; padding: 0; }
body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
${DOCUMENT_CONTENT_CSS}
${NOTES_CSS}
${PRINT_ONLY_CSS}
${extraCss}
</style>
</head>
<body>${wrapInContent ? `<div class="page__content">${bodyHtml}</div>` : bodyHtml}</body>
</html>`
}
