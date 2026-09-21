import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_NATURAL_LINE_HEIGHT,
  NATURAL_LINE_HEIGHTS,
} from '@services/document/line-metrics.js'

/**
 * A altura natural da linha, com os dois lados de verdade.
 *
 * O número é usado nas duas pontas: o leitor multiplica por ele no C#
 * (`BodyReader.LineHeightOf`) e a interface divide por ele aqui
 * (`paragraph-format.ts`). Duas tabelas que discordem fazem "1,5 linha" na tela
 * virar outra coisa no arquivo — o erro de 15 a 22 % que a conversão veio
 * corrigir —, e o pior é que nenhum teste de um lado só pegaria.
 *
 * Fica em `src/main` porque só aqui há Node: `src/services` é compilado também
 * para a web, e lá não existe `node:fs` para ler o arquivo do sidecar.
 */
describe('contrato da altura natural da linha', () => {
  it('a tabela é a mesma que o sidecar usa', () => {
    // O número tem de ser **um**: o leitor multiplica no C# e a interface divide
    // aqui. Duas tabelas que discordem fazem "1,5" na tela virar outra coisa no
    // arquivo, que é justamente o defeito que esta conversão veio corrigir.
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
