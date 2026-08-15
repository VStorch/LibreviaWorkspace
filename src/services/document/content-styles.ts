/**
 * Estilo do conteúdo do documento — **fonte única de verdade**.
 *
 * O mesmo texto é usado em dois lugares: injetado no editor e embutido no HTML
 * que gera o PDF. Se fossem duas folhas de estilo, elas divergiriam com o
 * tempo e o PDF deixaria de sair igual à tela — o risco registrado em
 * docs/00-plano-tecnico.md §6.3.
 *
 * Por isso não usa variáveis CSS do aplicativo: precisa ser autossuficiente
 * dentro de um documento HTML isolado.
 */
export const DOCUMENT_CONTENT_CSS = `
.page__content {
  outline: none;
  font-family: 'Liberation Serif', Georgia, serif;
  font-size: 12pt;
  line-height: 1.5;
  color: #111111;
}

.page__content > * + * { margin-top: 0.6em; }

.page__content h1 { font-size: 22pt; font-weight: 600; margin-top: 1em; }
.page__content h2 { font-size: 17pt; font-weight: 600; margin-top: 1em; }
.page__content h3 { font-size: 14pt; font-weight: 600; margin-top: 1em; }
.page__content h4 { font-size: 12pt; font-weight: 700; margin-top: 1em; }

.page__content ul,
.page__content ol { padding-left: 1.6em; }

.page__content li > p { margin: 0; }

.page__content a { color: #14538f; text-decoration: underline; }

.page__content img { max-width: 100%; height: auto; }

.page__content blockquote {
  border-left: 3px solid #c7ced6;
  padding-left: 1em;
  color: #444444;
}

.page__content table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
  margin: 0.8em 0;
}

.page__content th,
.page__content td {
  border: 1px solid #9aa3ad;
  padding: 4px 8px;
  vertical-align: top;
  position: relative;
}

.page__content th {
  background: #f0f2f4;
  font-weight: 600;
  text-align: left;
}
`

/**
 * Estilo só do editor: seleção de célula, alça de redimensionamento e a
 * representação visual da quebra de página. Nada disso existe no papel.
 */
export const EDITOR_ONLY_CSS = `
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

.page__content .page-break {
  border: none;
  border-top: 2px dashed #c7ced6;
  margin: 1.4em 0;
  position: relative;
  height: 0;
}

.page__content .page-break::after {
  content: 'Quebra de página';
  position: absolute;
  top: -0.75em;
  left: 50%;
  transform: translateX(-50%);
  background: #ffffff;
  padding: 0 8px;
  font-family: system-ui, sans-serif;
  font-size: 10px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #6b7280;
}

.page__content .ProseMirror-selectednode { outline: 2px solid #1f5fa9; }
`

/**
 * Ajustes que só valem no papel.
 *
 * A quebra de página deixa de ser uma linha tracejada e passa a ser uma quebra
 * de verdade; e evitamos os defeitos clássicos de paginação — título órfão no
 * pé da página, linha de tabela partida ao meio, imagem cortada.
 */
export const PRINT_ONLY_CSS = `
.page__content [data-page-break] {
  break-after: page;
  border: none;
  height: 0;
  margin: 0;
}

.page__content [data-page-break]::after { content: none; }

.page__content h1,
.page__content h2,
.page__content h3,
.page__content h4,
/* w:keepNext do documento: o bloco não fica sozinho no pé da página. É a
   mesma regra que as marcas de fim de página usam na tela — se divergissem, a
   marca cairia num lugar e o PDF quebraria noutro. */
.page__content [data-keep-next] { break-after: avoid; }

.page__content tr,
.page__content img { break-inside: avoid; }

.page__content thead { display: table-header-group; }
`
