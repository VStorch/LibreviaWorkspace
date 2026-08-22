/**
 * As fontes que viajam no instalador, e o nome pelo qual o documento as pede.
 *
 * Documento corporativo pede Calibri, Cambria, Arial, Times New Roman e Courier
 * New. Nenhuma existe num Linux limpo, e o que o Chromium põe no lugar tem
 * métrica própria: a linha quebra noutro ponto, e o documento de três páginas
 * vira quatro. Enquanto a paginação era só da exportação isso era um detalhe do
 * PDF; com a tela paginando, é o número que a pessoa lê na barra de status.
 *
 * As substitutas abaixo são **metricamente compatíveis** — cada glifo tem a
 * largura do glifo original, então a quebra cai no mesmo lugar. É o mesmo
 * conjunto que o LibreOffice usa para abrir documento do Word, e é o que torna
 * a comparação com ele uma comparação de verdade.
 *
 * | Do documento    | Empacotada        | Licença  |
 * | --------------- | ----------------- | -------- |
 * | Calibri         | Carlito           | OFL 1.1  |
 * | Cambria         | Caladea           | OFL 1.1  |
 * | Arial           | Liberation Sans   | OFL 1.1  |
 * | Times New Roman | Liberation Serif  | OFL 1.1  |
 * | Courier New     | Liberation Mono   | OFL 1.1  |
 *
 * O nome declarado no `@font-face` é o **da fonte original**. Assim um `w:rFonts
 * w:ascii="Calibri"` acha a substituta sem que ninguém precise reescrever o
 * documento — e se a máquina tiver a Calibri de verdade instalada, a pilha CSS
 * a prefere, porque o `local()` vem primeiro.
 */

/** O esquema é servido pelo processo main; ver `src/main/fonts.ts`. */
const SCHEME = 'librevia-font://fonts'

interface Substitute {
  /** Nome que o documento usa. */
  readonly declared: string
  /** Nome real da fonte instalada, quando existir na máquina. */
  readonly local: readonly string[]
  /** Prefixo dos arquivos empacotados. */
  readonly file: string
}

const SUBSTITUTES: readonly Substitute[] = [
  { declared: 'Calibri', local: ['Calibri', 'Carlito'], file: 'Carlito' },
  { declared: 'Cambria', local: ['Cambria', 'Caladea'], file: 'Caladea' },
  { declared: 'Arial', local: ['Arial', 'Liberation Sans'], file: 'LiberationSans' },
  { declared: 'Helvetica', local: ['Helvetica', 'Liberation Sans'], file: 'LiberationSans' },
  {
    declared: 'Times New Roman',
    local: ['Times New Roman', 'Liberation Serif'],
    file: 'LiberationSerif',
  },
  { declared: 'Courier New', local: ['Courier New', 'Liberation Mono'], file: 'LiberationMono' },
]

const FACES: readonly { suffix: string; weight: number; style: string }[] = [
  { suffix: 'Regular', weight: 400, style: 'normal' },
  { suffix: 'Bold', weight: 700, style: 'normal' },
  { suffix: 'Italic', weight: 400, style: 'italic' },
  { suffix: 'BoldItalic', weight: 700, style: 'italic' },
]

/**
 * `@font-face` para as cinco famílias, quatro variantes cada.
 *
 * `local()` antes de `url()` de propósito: numa máquina que tem a fonte
 * original, usá-la é mais fiel do que a substituta — e ainda evita carregar
 * arquivo que não precisa.
 *
 * `font-display: block` porque a medida depende da fonte. Com o padrão `auto` o
 * Chromium desenha com a fonte de reserva enquanto carrega, o documento é
 * medido com a métrica errada e as páginas se recontam sozinhas um instante
 * depois — a tela pisca e o número muda na frente de quem está lendo.
 */
export const DOCUMENT_FONT_CSS = SUBSTITUTES.flatMap((substitute) =>
  FACES.map(
    ({ suffix, weight, style }) => `@font-face {
  font-family: '${substitute.declared}';
  font-weight: ${weight};
  font-style: ${style};
  font-display: block;
  src: ${substitute.local.map((name) => `local('${name}')`).join(', ')},
       url('${SCHEME}/${substitute.file}-${suffix}.ttf') format('truetype');
}`,
  ),
).join('\n')

/** Os arquivos que precisam existir em `resources/fonts/`. */
export const BUNDLED_FONT_FILES: readonly string[] = [
  ...new Set(
    SUBSTITUTES.flatMap((substitute) => FACES.map(({ suffix }) => `${substitute.file}-${suffix}.ttf`)),
  ),
]
