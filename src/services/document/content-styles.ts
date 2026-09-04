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
     empacotada. Crase nenhuma aqui dentro: isto mora num template literal. */
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

.page__content img { max-width: 100%; height: auto; }

/*
  A imagem ancorada ao parágrafo é um bloco, e não uma palavra.

  Inline ela repousa sobre a linha de base e sobra por baixo a descida da
  fonte, que o Word não cobra: medido no LibreOffice, o parágrafo de uma
  captura ocupa a altura da captura. Com a descida, um documento de trinta
  capturas fecha uma folha depois — e as folhas passam a cortar em outro lugar.

  Medido no corpus: com esta regra o documento de evidências corta nas mesmas
  nove primeiras folhas que o LibreOffice, contra quatro sem ela.

  Pelo atributo, e não pela posição: o ProseMirror põe uma imagem vazia de
  serviço ao lado da de verdade para dar onde pôr o cursor, e com ela nenhuma
  imagem é filha única. A regra por posição valia no papel e não valia na tela,
  que é justamente como as duas divergem.

  Crase nenhuma aqui dentro: isto mora num template literal.
*/
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

.page__content img[data-anchored] { display: block; }

/*
  E a quebra de serviço do editor não abre linha depois dela.

  O ProseMirror põe um BR no fim do bloco para dar onde pôr o cursor. Com a
  imagem em bloco, esse BR cai numa linha própria e cobra a altura dela — na
  tela, que é quem pagina, e não no papel, que não tem BR nenhum. Eram 18 px
  por captura de diferença entre o que a tela mede e o que o papel imprime.
*/
.page__content img[data-anchored] ~ br,
.page__content img[data-anchored] ~ img.ProseMirror-separator { display: none !important; }

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
