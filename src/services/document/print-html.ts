import { DOCUMENT_CONTENT_CSS, PRINT_ONLY_CSS } from './content-styles.js'

/**
 * Monta o HTML que o Chromium vai transformar em PDF.
 *
 * O corpo vem do próprio editor (`editor.getHTML()`), então o que é impresso é
 * literalmente o que foi editado — não uma segunda renderização a partir do
 * modelo, que poderia divergir.
 *
 * Não há margem nem largura fixada aqui: quem define tamanho de página e
 * margens é o `printToPDF` (ver @services/pdf/page-setup.ts). Duplicar isso no
 * CSS produziria margem dobrada.
 */
export function buildPrintHtml(bodyHtml: string, title: string): string {
  return `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
html, body { margin: 0; padding: 0; }
body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
${DOCUMENT_CONTENT_CSS}
${PRINT_ONLY_CSS}
</style>
</head>
<body><div class="page__content">${bodyHtml}</div></body>
</html>`
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
