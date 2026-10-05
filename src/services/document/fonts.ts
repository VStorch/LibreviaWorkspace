/**
 * **Metric-compatible** substitutes: each glyph has the original's width, so lines break in the
 * same place, as in LibreOffice.
 *
 * | Document        | Bundled           | License  |
 * | --------------- | ----------------- | -------- |
 * | Calibri         | Carlito           | OFL 1.1  |
 * | Cambria         | Caladea           | OFL 1.1  |
 * | Arial           | Liberation Sans   | OFL 1.1  |
 * | Times New Roman | Liberation Serif  | OFL 1.1  |
 * | Courier New     | Liberation Mono   | OFL 1.1  |
 *
 * `@font-face` uses the **original** font's name, so `w:ascii="Calibri"` finds the substitute, and
 * an installed Calibri wins through `local()`.
 */

/** Main serves the scheme; see `src/main/fonts.ts`. */
const SCHEME = 'librevia-font://fonts'

interface Substitute {
  readonly declared: string
  /** Installed families that will do, when present. */
  readonly local: readonly string[]
  /** Prefix of the bundled files. */
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
 * `local()` matches by font name, not family: `local('Liberation Sans')` in a bold rule would serve
 * the regular face. Full name and PostScript name, because both occur.
 */
function localNames(family: string, words: string): string[] {
  if (words === '') return [family]
  return [`${family} ${words}`, `${family.replaceAll(' ', '')}-${words.replaceAll(' ', '')}`]
}

/**
 * `font-display: block` because measuring depends on the font: with `auto`, the document would be
 * measured with the fallback font and repaginated a moment later.
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

/** The files that must exist in `resources/fonts/`. */
export const BUNDLED_FONT_FILES: readonly string[] = [
  ...new Set(
    SUBSTITUTES.flatMap((substitute) => FACES.map(({ suffix }) => `${substitute.file}-${suffix}.ttf`)),
  ),
]
