import { DOCUMENT_CONTENT_CSS, NOTES_CSS, PRINT_ONLY_CSS } from './content-styles.js'

/**
 * O corpo vem do próprio editor (`editor.getHTML()`). Tamanho de página e
 * margens são do `printToPDF` (`@services/pdf/page-setup.ts`). `extraCss` é para
 * a planilha.
 */
export function buildPrintHtml(
  bodyHtml: string,
  title: string,
  extraCss = '',
  /** O documento paginado **não** se envolve: cada folha traz o seu `.page__content`. */
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

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
