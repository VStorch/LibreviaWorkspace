/**
 * Substitutas **metricamente compatíveis**: cada glifo tem a largura do
 * original, então a linha quebra no mesmo lugar, como no LibreOffice.
 *
 * | Do documento    | Empacotada        | Licença  |
 * | --------------- | ----------------- | -------- |
 * | Calibri         | Carlito           | OFL 1.1  |
 * | Cambria         | Caladea           | OFL 1.1  |
 * | Arial           | Liberation Sans   | OFL 1.1  |
 * | Times New Roman | Liberation Serif  | OFL 1.1  |
 * | Courier New     | Liberation Mono   | OFL 1.1  |
 *
 * O `@font-face` usa o nome **da fonte original**, então `w:ascii="Calibri"`
 * acha a substituta, e a Calibri instalada vence pelo `local()`.
 */

/** O esquema é servido pelo processo main; ver `src/main/fonts.ts`. */
const SCHEME = 'librevia-font://fonts'

interface Substitute {
  readonly declared: string
  /** Famílias instaladas que servem, quando existirem na máquina. */
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

const FACES: readonly { suffix: string; weight: number; style: string; words: string }[] = [
  { suffix: 'Regular', weight: 400, style: 'normal', words: '' },
  { suffix: 'Bold', weight: 700, style: 'normal', words: 'Bold' },
  { suffix: 'Italic', weight: 400, style: 'italic', words: 'Italic' },
  { suffix: 'BoldItalic', weight: 700, style: 'italic', words: 'Bold Italic' },
]

/**
 * `local()` casa por nome de fonte, e não de família: `local('Liberation Sans')`
 * numa regra de negrito serviria o corte normal. Nome cheio e PostScript, porque
 * os dois aparecem.
 */
function localNames(family: string, words: string): string[] {
  if (words === '') return [family]
  return [`${family} ${words}`, `${family.replaceAll(' ', '')}-${words.replaceAll(' ', '')}`]
}

/**
 * `font-display: block` porque a medida depende da fonte: com `auto`, o
 * documento seria medido com a fonte de reserva e repaginado um instante depois.
 */
export const DOCUMENT_FONT_CSS = SUBSTITUTES.flatMap((substitute) =>
  FACES.map(({ suffix, weight, style, words }) => {
    const locals = substitute.local.flatMap((family) => localNames(family, words))

    return `@font-face {
  font-family: '${substitute.declared}';
  font-weight: ${weight};
  font-style: ${style};
  font-display: block;
  src: ${locals.map((name) => `local('${name}')`).join(', ')},
       url('${SCHEME}/${substitute.file}-${suffix}.ttf') format('truetype');
}`
  }),
).join('\n')

/** Os arquivos que precisam existir em `resources/fonts/`. */
export const BUNDLED_FONT_FILES: readonly string[] = [
  ...new Set(
    SUBSTITUTES.flatMap((substitute) => FACES.map(({ suffix }) => `${substitute.file}-${suffix}.ttf`)),
  ),
]
