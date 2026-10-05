import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_NATURAL_LINE_HEIGHT,
  NATURAL_LINE_HEIGHTS,
} from '@services/document/line-metrics.js'

/**
 * The natural line height, with both real sides: the reader multiplies by it in C#
 * (`BodyReader.LineHeightOf`) and the UI divides (`paragraph-format.ts`). Lives in `src/main`, the
 * only place with `node:fs`.
 */
describe('contrato da altura natural da linha', () => {
  it('a tabela é a mesma que o sidecar usa', () => {
    // The number must be **one**: the reader multiplies in C# and the UI divides here. Two tables
    // that disagree turn "1.5" on screen into something else in the file.
    const source = readFileSync(
      new URL('../../../sidecar/src/Librevia.Format/Docx/LineMetrics.cs', import.meta.url),
      'utf8',
    )

    const declared = new Map<string, number>()
    for (const [, font, value] of source.matchAll(/\["([^"]+)"\]\s*=\s*([\d.]+|LiberationSerif)/g)) {
      declared.set(
        font!.toLowerCase(),
        value === 'LiberationSerif' ? DEFAULT_NATURAL_LINE_HEIGHT : Number(value),
      )
    }

    expect(declared.size).toBeGreaterThan(0)
    expect(Object.fromEntries(declared)).toEqual(NATURAL_LINE_HEIGHTS)
  })

})
