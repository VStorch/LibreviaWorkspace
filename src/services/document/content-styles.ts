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
import { DOCUMENT_FONT_CSS } from './fonts.js'

export const DOCUMENT_CONTENT_CSS = `
${DOCUMENT_FONT_CSS}

.page__content {
  outline: none;
  /*
    Coluna de caixas flexíveis, e não fluxo comum, por uma razão só: **as
    margens não podem se juntar**. O CSS funde a margem de baixo de um bloco com
    a de cima do seguinte e fica com a maior; o Word e o LibreOffice somam as
    duas. Num documento em que cada parágrafo pede 14 pt depois e o seguinte
    14 pt antes, a diferença é meia linha por junta — e ela se acumula até a
    folha cortar noutro lugar.
  */
  display: flex;
  flex-direction: column;
  /* Pelo nome do documento: a regra @font-face acima resolve para a
     empacotada. Crase nenhuma aqui dentro: isto mora num template literal.
     No documento, a fonte, o tamanho e a entrelinha daqui (e os títulos logo
     abaixo) são sobrepostos pelo CSS dos estilos (style-css.ts); valem sozinhos
     só fora dele, na planilha impressa. */
  font-family: 'Times New Roman', 'Liberation Serif', Georgia, serif;
  font-size: 12pt;
  line-height: 1.5;
  color: #111111;
}

.page__content > * + * { margin-top: 0.6em; }

/*
  O parágrafo vazio também ocupa uma linha.

  No editor ele já ocupa: o ProseMirror põe um <br> invisível dentro para dar
  onde pôr o cursor. No papel não há cursor, o serializador emite <p></p> e um
  bloco sem conteúdo tem altura zero. As folhas eram recortadas pela medida da
  tela e impressas sem essas linhas, então o papel subia o texto todo e a
  primeira linha ia parar debaixo do cabeçalho. Este pedaço vazio de linha
  devolve a altura sem devolver tinta.
*/
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

/**
 * Sobrescrito e subscrito sem esticar a linha.
 *
 * O padrão do navegador — \`vertical-align: super\` com \`font-size: smaller\` — sobe
 * o glifo e **cresce a caixa da linha** com ele: a linha que tem um expoente fica
 * mais alta que as vizinhas, e num documento paginado ao vivo isso desloca a
 * quebra de página. O Word não faz isso. \`line-height: 0\` devolve a medida da
 * linha ao texto normal, que é o que a paginação precisa medir; o tamanho é
 * declarado em fração de propósito, porque \`smaller\` encolhe de novo a cada
 * expoente dentro de outro.
 */
.page__content sup,
.page__content sub { font-size: 0.65em; line-height: 0; }

.page__content img { max-width: 100%; height: auto; }

/* Imagem sem parágrafo, alinhada por atributo próprio. O editor de hoje não a
   produz — a imagem é conteúdo de linha, e quem a alinha é o parágrafo —, mas
   o que foi salvo quando ela era um bloco ainda imprime assim. */
.page__content img[data-align='center'] { display: block; margin-inline: auto; }
.page__content img[data-align='right'] { display: block; margin-left: auto; }

/*
  A marca da lista é a que o documento declara, à distância que ele pede.

  O CSS escolheria a bolinha e a encostaria no texto; o documento diz o
  caractere em w:lvlText e a distância em w:ind/@hanging. Desenhada por um
  pseudo-elemento porque o marcador nativo não se posiciona — e é justamente a
  distância que faz o recuo pendente do Word.
*/
.page__content ul[data-marker] { list-style: none; }

.page__content ul[data-marker] > li > :first-child::before {
  content: var(--marca);
  display: inline-block;
  width: var(--pendente, 1em);
  margin-left: calc(-1 * var(--pendente, 1em));
}

/*
  A imagem ancorada ao parágrafo é um bloco, e não uma palavra.

  Inline ela repousa sobre a linha de base e sobra por baixo a descida da
  fonte, que o Word não cobra: medido no LibreOffice, o parágrafo de uma
  captura ocupa a altura da captura. Com a descida, um documento de trinta
  capturas fecha uma folha depois — e as folhas passam a cortar em outro lugar.

  Pelo atributo, e não pela posição: o ProseMirror põe uma imagem vazia de
  serviço ao lado da de verdade para dar onde pôr o cursor, e com ela nenhuma
  imagem é filha única. A regra por posição valia no papel e não valia na tela,
  que é justamente como as duas divergem.

  E ela ocupa a coluna, não a caixa do parágrafo: no Word a ancorada não é
  texto, posiciona-se pela coluna, e o recuo do parágrafo que a carrega não a
  estreita. Sem descontar o recuo, uma captura de 140 mm dentro de um parágrafo
  recuado 12,7 mm era espremida — e como a altura segue a largura, encurtava
  junto, e a legenda seguinte passava a caber numa folha em que o LibreOffice já
  não a punha.

  Crase nenhuma aqui dentro: isto mora num template literal.
*/
/* Na tela quem fica no lugar da imagem, como filho do parágrafo, é o embrulho
   do NodeView (.node-image), e a moldura das alças mora dentro dele. A regra
   tem de pegar o embrulho: aplicada à moldura, que não é filha do parágrafo,
   as duas regras de baixo não casavam com nada — e a paginação da tela só batia
   com a do papel por acaso. A moldura encolhe à imagem para as alças ficarem
   nos cantos dela, e não nos da coluna. */
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

/*
  E a quebra de serviço do editor não abre linha depois dela.

  O ProseMirror põe um BR no fim do bloco para dar onde pôr o cursor. Com a
  imagem em bloco, esse BR cai numa linha própria e cobra a altura dela — na
  tela, que é quem pagina, e não no papel, que não tem BR nenhum. Eram 18 px
  por captura de diferença entre o que a tela mede e o que o papel imprime.
*/
.page__content img[data-anchored] ~ br,
.page__content img[data-anchored] ~ img.ProseMirror-separator,
.page__content .node-image[data-anchored] ~ br,
.page__content .node-image[data-anchored] ~ img.ProseMirror-separator { display: none !important; }

/*
  Mas o parágrafo dela tem uma linha, e a linha ocupa lugar.

  No Word a captura ancorada é um quadro que flutua: o parágrafo continua sendo
  um parágrafo, com a linha vazia dele, e o quadro empurra o que vem depois. Com
  a captura ocupando a largura da coluna inteira, não sobra onde a linha caber ao
  lado, e ela vai para baixo.

  Medido no LibreOffice, no documento de evidências do corpus: entre duas
  capturas encostadas uma na outra ele deixa 12 pt, que é exatamente a entrelinha
  do parágrafo. Nós as encostávamos, e o erro se somava a cada captura — num
  documento de trinta, folhas inteiras de diferença.

  Isto foi medido antes e concluído ao contrário: a linha parecia sobrar. A
  medida estava certa e a base, errada — a entrelinha saía 1,15 vez curta e a
  captura 2% estreita, e os dois erros escondiam este. Com eles corrigidos, três
  dos quatro documentos de evidências passam a cortar nas mesmas folhas.

  Dentro do parágrafo, e não como margin-bottom. A margem parece mais natural e
  chega a consertar uma folha do documento de quinze — no LibreOffice esta linha
  pode passar para a folha seguinte enquanto o quadro fica, e a margem imita isso
  porque não conta no pé da folha. Mas medido nos quatro documentos ela custa
  caro: 5/5, 5/5 e 5/5 folhas iguais caem para 3/5, 1/5 e 0/5. Dentro do
  parágrafo a linha conta sempre, e é o que mais se parece com o resultado.

  Crase nenhuma aqui dentro: isto mora num template literal.
*/
.page__content p:has(> img[data-anchored])::after,
.page__content p:has(> .node-image[data-anchored])::after {
  content: '';
  display: block;
  height: 1lh;
}

/*
  A marca de seção não ocupa linha.

  No OOXML a seção termina num w:sectPr guardado dentro do w:pPr de um parágrafo
  vazio: o parágrafo é a marca, e não uma linha de texto. O LibreOffice, que é
  quem grava documentos assim, não lhe dá altura — o documento de evidências do
  corpus tem sete seções de mesma geometria e seis marcas no meio do texto, e
  cada uma valia uma linha aqui.

  Continua no documento, e não some: é ela que a gravação devolve ao arquivo.
*/
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
/**
 * O conteúdo no tema escuro.
 *
 * Mora **dentro** de `EDITOR_ONLY_CSS` de propósito, e é a decisão mais
 * importante deste arquivo desde que ele existe: o tema é como a tela desenha,
 * e o papel impresso é branco em qualquer tema. Uma regra escura em
 * `DOCUMENT_CONTENT_CSS` sairia no PDF, e o primeiro PDF preto exportado por
 * quem estava lendo no escuro seria um defeito difícil de rastrear até aqui.
 * `print-html.ts` monta o papel com `DOCUMENT_CONTENT_CSS + PRINT_ONLY_CSS`, e
 * nunca com este.
 *
 * ## O que troca de cor, e o que não
 *
 * Só troca o que o **documento não pediu**. `color: #111111` lá em cima não é
 * uma escolha do autor: é o preto padrão de quem não declarou cor nenhuma, e
 * padrão acompanha o tema. Já o que o autor pintou chega como estilo em linha,
 * que ganha de qualquer seletor daqui — então continua exatamente como estava,
 * sem uma linha de código para garanti-lo.
 *
 * Vale para a borda da tabela, o fundo do cabeçalho de tabela e a barra da
 * citação pelo mesmo motivo: são padrões nossos, não do arquivo.
 *
 * ## O caso que isto não resolve
 *
 * Um documento que declara o texto como preto — `w:color w:val="000000"` — fica
 * preto sobre papel escuro, e não se lê. É o preço de honrar a cor do autor:
 * em linha, ela ganha, e distinguir "o autor quis preto" de "o Word escreveu
 * preto por escrever" exigiria mexer no que `getHTML()` produz — que é
 * exatamente o que alimenta a gravação cirúrgica. Trocar legibilidade de um
 * caso incomum por risco no que o projeto existe para proteger não vale a
 * troca. Word e LibreOffice só gravam `w:color` quando alguém mudou a cor, de
 * modo que o corpo de um documento comum não declara nenhuma e acompanha o
 * tema normalmente.
 */
const DARK_CONTENT_CSS = `
:root[data-theme='dark'] .page__content {
  color: var(--text);
}

/* O papel. A mesma cor do resto das superfícies, para o documento não parecer
   uma janela recortada dentro da casca. */
:root[data-theme='dark'] .paper {
  background: var(--surface);
}

/* Azul de link clareado pelo mesmo motivo que a ênfase: #14538f sobre papel
   escuro dá 2.1:1. */
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

export const EDITOR_ONLY_CSS = `
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

/*
  A quebra de página não desenha nada.

  Ela era uma linha tracejada escrita "QUEBRA DE PÁGINA" — a marca fazia sentido
  quando a tela era uma tira contínua e a quebra não tinha efeito nenhum de se
  ver. Agora a folha termina ali de verdade, e é isso que o Word e o LibreOffice
  mostram na vista de impressão: nada. Na capa do modelo de manual a marca ainda
  caía no meio do desenho, porque o título e o subtítulo moram em caixas
  posicionadas e o fluxo ali é quase vazio.

  Altura zero, sem margem: ela também não pode ocupar lugar na folha. Continua
  selecionável — o contorno de nó selecionado a mostra — e continua sendo o que
  se apaga com Backspace no começo da folha seguinte, como no Word.
*/
.page__content .page-break {
  border: none;
  margin: 0;
  height: 0;
}

.page__content .ProseMirror-selectednode { outline: 2px solid #1f5fa9; }

/*
  A moldura das alças de redimensionamento.

  Em linha, como a imagem que ela envolve, e com line-height zero: sem isto a
  linha de dentro da moldura cobraria a descida da fonte por baixo da imagem, e
  a paginação mediria cada imagem alguns pixels mais alta do que o papel a
  imprime.
*/
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

/*
  As marcas de formatação não podem mudar a medida da linha.

  É a mesma lição do sobrescrito, e aqui ela custaria mais caro: a marca de
  parágrafo aparece no fim de **todo** parágrafo, então um pixel de altura a mais
  se soma em cada linha e a paginação inteira desliza. O estilo que vem com a
  extensão oficial desenha as marcas com line-height 1em — daí o injectCSS: false
  em editor-extensions.ts e este pedaço aqui.

  Largura e altura zero, entrelinha zero: a caixa não ocupa lugar nenhum. O glifo
  continua visível porque overflow é visível por padrão, e pousa na linha de base
  porque é a linha de base interna que um inline-block de overflow visível
  apresenta ao redor.

  Só no editor: o papel não mostra marca de formatação, como no Word. Elas são
  decoração do ProseMirror e por isso nem chegam ao HTML que gera o PDF.

  Crase nenhuma aqui dentro: isto mora num template literal.
*/
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

/* A imagem de serviço que o ProseMirror põe ao lado da marca também não cobra
   altura — é o mesmo cuidado que a imagem ancorada pediu mais acima. */
.page__content .tiptap-invisible-character + img.ProseMirror-separator {
  width: 0 !important;
  height: 0 !important;
  pointer-events: none;
  user-select: none;
}
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
