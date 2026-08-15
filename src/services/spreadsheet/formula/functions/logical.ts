/**
 * Funções lógicas.
 *
 * `SE` e `SEERRO` não estão aqui: elas precisam receber os argumentos sem
 * avaliar e vivem no avaliador, junto da explicação do porquê.
 */

import type { Argument } from '../evaluate.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { toBoolean, type Scalar } from '../values.js'
import { VARIADIC, define, single, valuesIn, type FunctionDefinition } from './kit.js'

/**
 * `E` e `OU` **não** interrompem no primeiro resultado.
 *
 * O Excel avalia todos os argumentos, e o erro de qualquer um contamina o
 * resultado. Parar cedo faria `=E(FALSO;1/0)` devolver FALSO aqui e `#DIV/0!`
 * no Excel, para a mesma planilha.
 */
function fold(
  args: readonly Argument[],
  combine: (a: boolean, b: boolean) => boolean,
  seed: boolean,
): Scalar {
  let result = seed
  let seen = false

  for (const value of valuesIn(args)) {
    // Vazio dentro de intervalo é ignorado, como no Excel.
    if (value === null) continue

    const flag = toBoolean(value)
    if (isFormulaError(flag)) return flag
    result = combine(result, flag)
    seen = true
  }

  // Nenhum valor lógico entre os argumentos: não há o que responder.
  return seen ? result : FormulaError.Value
}

export const LOGICAL: readonly FunctionDefinition[] = [
  define(['E', 'AND'], 1, VARIADIC, (args) => fold(args, (a, b) => a && b, true)),
  define(['OU', 'OR'], 1, VARIADIC, (args) => fold(args, (a, b) => a || b, false)),

  define(['NÃO', 'NAO', 'NOT'], 1, 1, (args) => {
    const flag = toBoolean(single(args[0]))
    return isFormulaError(flag) ? flag : !flag
  }),

  // As quatro examinam o argumento em vez de usá-lo, então recebem o erro.
  define(['ÉERROS', 'EERROS', 'ISERROR'], 1, 1, (args) => isFormulaError(single(args[0])), true),
  define(['É.NÃO.DISP', 'ENAODISP', 'ISNA'], 1, 1, (args) => single(args[0]) === FormulaError.NA, true),
  define(['ÉNÚM', 'ENUM', 'ISNUMBER'], 1, 1, (args) => typeof single(args[0]) === 'number', true),
  define(['ÉTEXTO', 'ETEXTO', 'ISTEXT'], 1, 1, (args) => typeof single(args[0]) === 'string', true),
  define(['ÉCÉL.VAZIA', 'ECELVAZIA', 'ISBLANK'], 1, 1, (args) => single(args[0]) === null, true),
]
