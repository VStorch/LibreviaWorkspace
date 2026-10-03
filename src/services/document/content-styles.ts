/**
 * O mesmo CSS vai ao editor e ao HTML que gera o PDF, para o papel sair igual à
 * tela. Não usa variáveis CSS do aplicativo: precisa valer num HTML isolado.
 */
import { DOCUMENT_FONT_CSS } from './fonts.js'

export const DOCUMENT_CONTENT_CSS = `
${DOCUMENT_FONT_CSS}

.page__content {
  outline: none;
  /* Flex, e não fluxo comum, para as margens não se fundirem: o CSS fica com a
     maior, e o Word e o LibreOffice somam as duas. */
  display: flex;
  flex-direction: column;
  /* No documento, o CSS dos estilos (style-css.ts) sobrepõe isto; vale sozinho
     na planilha impressa. */
  font-family: 'Times New Roman', 'Liberation Serif', Georgia, serif;
  font-size: 12pt;
  line-height: 1.5;
  color: #111111;
}

.page__content > * + * { margin-top: 0.6em; }

/* O parágrafo vazio ocupa uma linha: no papel não há o <br> que o ProseMirror
   põe para o cursor, e sem isto o texto subiria. */
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

/* O sumário: o estilo "toc N" do Word pede tabulação à direita com pontinhos.
   Sem paradas de tabulação no editor, a entrada vira uma fileira flexível. O
   link da entrada não se pinta, como no Word. */
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

/* A fonte de matemática é a do sistema: \`math\` é a família genérica que o
   Chromium resolve para a fonte com tabela OpenType MATH que houver. A caixa
   (\`m:borderBox\`) é um \`mrow\` com borda, porque o MathML Core não tem
   \`menclose\`. */
.page__content .equacao math {
  font-family: 'STIX Two Math', 'Cambria Math', 'Noto Sans Math', 'DejaVu Math TeX Gyre', math;
}
.page__content .equacao--exibicao { display: block; text-align: center; break-inside: avoid; }
.page__content .equacao--exibicao[data-jc='left'] { text-align: left; }
.page__content .equacao--exibicao[data-jc='right'] { text-align: right; }
.page__content .equacao--exibicao math { display: inline math; math-style: normal; }
.page__content .equacao .omml-caixa { border: 1px solid currentColor; padding: 0.1em; }

/* Sobrescrito e subscrito sem esticar a linha, como no Word: \`line-height: 0\`
   devolve a medida da linha, e o tamanho em fração não encolhe de novo a cada
   expoente aninhado. */
.page__content sup,
.page__content sub { font-size: 0.65em; line-height: 0; }

.page__content img { max-width: 100%; height: auto; }

/* Imagem como bloco, alinhada por atributo: só vem de rascunhos antigos. */
.page__content img[data-align='center'] { display: block; margin-inline: auto; }
.page__content img[data-align='right'] { display: block; margin-left: auto; }

/* A marca da lista vem pronta em --lista-marca (list-numbering.ts), contada
   como o Word conta, e é desenhada por pseudo-elemento porque o marcador nativo
   não se posiciona no recuo pendente (w:ind/@hanging). O recuo é relativo ao da
   lista de fora (--lista-recuo); !important porque o nó pode trazer o absoluto
   em estilo inline. */
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
  /* Marca maior que o recuo pendente: o Word salta para a tabulação seguinte;
     aqui, ao menos meia letra de ar. */
  padding-right: 0.5em;
  text-indent: 0;
  white-space: pre;
}

/* A imagem ancorada é um bloco do tamanho da coluna, como no Word: inline ela
   cobraria a descida da fonte, e o recuo do parágrafo não a estreita. Pelo
   atributo, e não pela posição, porque o ProseMirror põe uma imagem de serviço
   ao lado. Na tela a regra pega o embrulho do NodeView (.node-image), filho do
   parágrafo; a moldura encolhe à imagem para as alças ficarem nos cantos. */
.page__content img[data-anchored],
.page__content .node-image[data-anchored] {
  display: block;
  margin-left: calc(-1 * var(--recuo, 0mm));
  max-width: calc(100% + var(--recuo, 0mm) + var(--recuo-direita, 0mm));
}

.page__content .node-image[data-anchored] > .image-frame {
  display: block;
  width: fit-content;
}

/* O <br> de serviço do ProseMirror não abre linha depois da imagem em bloco:
   o papel não tem esse <br>. */
.page__content img[data-anchored] ~ br,
.page__content img[data-anchored] ~ img.ProseMirror-separator,
.page__content .node-image[data-anchored] ~ br,
.page__content .node-image[data-anchored] ~ img.ProseMirror-separator { display: none !important; }

/* O parágrafo da imagem ancorada mantém a linha dele por baixo da imagem,
   dentro do parágrafo: é o que o LibreOffice mede entre duas capturas. */
.page__content p:not([data-anchor-text]):has(> img[data-anchored])::after,
.page__content p:not([data-anchor-text]):has(> .node-image[data-anchored])::after {
  content: '';
  display: block;
  height: 1lh;
}


/* A marca de seção (o parágrafo vazio que guarda o w:sectPr) não ocupa linha,
   como no LibreOffice. Continua no documento: a gravação a devolve. */
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
  /* Sem margem própria, como no Word: o ar em volta é o dos parágrafos vizinhos. */
  margin: 0;
}

/* A margem de célula vem da tabela (--cell-margins); sem ela, 60/120 twips, o
   TableNormal do Word. */
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

/**
 * O tema escuro mora aqui, e nunca em `DOCUMENT_CONTENT_CSS`: o papel impresso é
 * branco em qualquer tema. Só troca de cor o que o documento não pediu; a cor
 * do autor vem em estilo inline e ganha. Um texto declarado preto fica preto
 * sobre o papel escuro: distingui-lo exigiria mexer no HTML que alimenta a
 * gravação cirúrgica.
 */
const DARK_CONTENT_CSS = `
:root[data-theme='dark'] .page__content {
  color: var(--text);
}

/* A mesma cor das outras superfícies, para o papel não parecer recortado. */
:root[data-theme='dark'] .paper {
  background: var(--surface);
}

/* #14538f sobre papel escuro dá 2.1:1. */
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

/* A altura do separador vem da paginação; o traço fica no meio dele. */
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
/* clip, e não hidden: a área não rola com o cursor. flow-root porque clip não
   isola a margem negativa da continuação. */
.paper-notes__slot { overflow: clip; display: flow-root; }
.note-body { display: flex; flex-direction: column; }
.note-number { vertical-align: super; font-size: 0.65em; line-height: 0; }
`

/** Só do editor: seleção de célula, alça de redimensionamento, quebra de página. */
export const EDITOR_ONLY_CSS = `
/* O sombreado cinza do Word mostra que o número é calculado. Só na tela. */
.page__content .field.ProseMirror-selectednode { background: #d9d9d9; outline: none; }

/* O número é decoração, pela ordem no texto (note-ref.ts); no papel vem escrito. */
.page__content .note-ref::after { content: attr(data-note-number); }
.page__content .note-ref.ProseMirror-selectednode { background: #d9d9d9; outline: none; }

/* A equação travada, com construção que a tela não desenha, leva um traço por
   baixo, e a dica diz o que falta. */
.page__content .equacao.ProseMirror-selectednode { background: #d9d9d9; outline: none; }
.page__content .equacao--travada { text-decoration: underline dotted #b0b0b0; }

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

/* A quebra de página não desenha nada, como na vista de impressão do Word e do
   LibreOffice: a folha termina ali. Continua selecionável e apagável com
   Backspace no começo da folha seguinte. */
.page__content .page-break {
  border: none;
  margin: 0;
  height: 0;
}

.page__content .ProseMirror-selectednode { outline: 2px solid #1f5fa9; }

/* Em linha e sem entrelinha: sem isto a moldura cobraria a descida da fonte, e
   a tela mediria a imagem mais alta que o papel. */
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

/* As marcas de formatação não mudam a medida da linha: largura, altura e
   entrelinha zero, com o glifo visível pelo overflow. Por isso injectCSS: false
   em editor-extensions.ts. */
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

/* A imagem de serviço ao lado da marca também não cobra altura. */
.page__content .tiptap-invisible-character + img.ProseMirror-separator {
  width: 0 !important;
  height: 0 !important;
  pointer-events: none;
  user-select: none;
}
`

/** Só no papel: título órfão, linha de tabela partida e imagem cortada. */
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
/* A mesma regra das marcas de fim de página da tela. */
.page__content [data-keep-next] { break-after: avoid; }

.page__content tr,
.page__content img { break-inside: avoid; }

.page__content thead { display: table-header-group; }

/* Como na tela (styles.css); nos outros modos a impressão sai sem as marcas
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
