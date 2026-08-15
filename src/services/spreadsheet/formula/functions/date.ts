/**
 * Funções de data.
 *
 * Datas são números de série contados de 1899-12-30, herança do Lotus 1-2-3 —
 * a explicação inteira está em `format.ts`, que é quem converte. Aqui só se
 * fazem contas com o número, o que é justamente a vantagem do formato: a
 * diferença entre duas datas é uma subtração.
 */

import { dateToSerial, serialToDate } from '../../format.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { define, numberArg, type FunctionDefinition } from './kit.js'

const MS_PER_DAY = 86_400_000

/** Parte do dia já decorrida, que é a parte fracionária do número de série. */
function timeFraction(date: Date): number {
  const midnight = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime()
  return (date.getTime() - midnight) / MS_PER_DAY
}

function partOf(serial: number | FormulaError, pick: (date: Date) => number): number | FormulaError {
  if (isFormulaError(serial)) return serial
  if (serial < 0) return FormulaError.Num
  return pick(serialToDate(Math.floor(serial)))
}

export const DATE: readonly FunctionDefinition[] = [
  define(['HOJE', 'TODAY'], 0, 0, (_args, context) => dateToSerial(context.now())),

  define(['AGORA', 'NOW'], 0, 0, (_args, context) => {
    const now = context.now()
    return dateToSerial(now) + timeFraction(now)
  }),

  define(['DATA', 'DATE'], 3, 3, (args) => {
    const year = numberArg(args[0])
    if (isFormulaError(year)) return year
    const month = numberArg(args[1])
    if (isFormulaError(month)) return month
    const day = numberArg(args[2])
    if (isFormulaError(day)) return day

    // Mês treze é janeiro do ano seguinte: o Date do JavaScript já normaliza,
    // e o Excel faz o mesmo.
    const date = new Date(Math.trunc(year), Math.trunc(month) - 1, Math.trunc(day))
    const serial = dateToSerial(date)
    return serial < 0 ? FormulaError.Num : serial
  }),

  // getUTCFullYear e não getFullYear: serialToDate monta a data em UTC, e ler
  // em fuso local devolveria o dia anterior a oeste de Greenwich.
  define(['ANO', 'YEAR'], 1, 1, (args) => partOf(numberArg(args[0]), (date) => date.getUTCFullYear())),
  define(['MÊS', 'MES', 'MONTH'], 1, 1, (args) =>
    partOf(numberArg(args[0]), (date) => date.getUTCMonth() + 1),
  ),
  define(['DIA', 'DAY'], 1, 1, (args) => partOf(numberArg(args[0]), (date) => date.getUTCDate())),

  /** Dia da semana, com domingo valendo um — o padrão do Excel. */
  define(['DIA.DA.SEMANA', 'WEEKDAY'], 1, 2, (args) => {
    const day = partOf(numberArg(args[0]), (date) => date.getUTCDay())
    if (isFormulaError(day)) return day

    const kind = args.length > 1 ? numberArg(args[1]) : 1
    if (isFormulaError(kind)) return kind

    if (kind === 1) return day + 1
    if (kind === 2) return day === 0 ? 7 : day
    if (kind === 3) return day === 0 ? 6 : day - 1
    return FormulaError.Num
  }),
]
