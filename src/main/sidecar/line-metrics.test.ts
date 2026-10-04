import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_NATURAL_LINE_HEIGHT,
  NATURAL_LINE_HEIGHTS,
} from '@services/document/line-metrics.js'

/**
 * A altura natural da linha, com os dois lados de verdade: o leitor multiplica por
 * ela no C# (`BodyReader.LineHeightOf`) e a interface divide (`paragraph-format.ts`).
 * Fica em `src/main`, o único lugar com `node:fs`.
 */
describe('contrato da altura natural da linha', () => {
  it('a tabela é a mesma que o sidecar usa', () => {
    // O número tem de ser **um**: o leitor multiplica no C# e a interface
    // divide aqui. Duas tabelas que discordem fazem "1,5" na tela virar outra
    // coisa no arquivo.
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
