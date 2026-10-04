/** Formatar na hora de exibir, e não ao digitar, como o XLSX: somar moeda não tropeça no "R$". */

import { CellFormat, type Cell, type CellStyle, type CellValue } from './model.js'

const LOCALE = 'pt-BR'

export function formatCell(cell: Cell | undefined): string {
  if (cell === undefined || cell.value === undefined) return ''

  const { value } = cell
  if (typeof value === 'boolean') return value ? 'VERDADEIRO' : 'FALSO'

  const format = cell.style?.format ?? CellFormat.General
  if (format === CellFormat.Text) return String(value)

  const numeric = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(numeric)) return String(value)

  return formatNumeric(value, numeric, format, cell.style?.decimals)
}

function formatNumeric(
  value: CellValue,
  numeric: number,
  format: CellFormat,
  decimals: number | undefined,
): string {
  switch (format) {
    case CellFormat.Currency:
      return numeric.toLocaleString(LOCALE, {
        style: 'currency',
        currency: 'BRL',
        minimumFractionDigits: decimals ?? 2,
        maximumFractionDigits: decimals ?? 2,
      })

    case CellFormat.Percent:
      return numeric.toLocaleString(LOCALE, {
        style: 'percent',
        minimumFractionDigits: decimals ?? 0,
        maximumFractionDigits: decimals ?? 2,
      })

    case CellFormat.Date:
      return formatDate(numeric)

    case CellFormat.Number:
      return numeric.toLocaleString(LOCALE, {
        minimumFractionDigits: decimals ?? 0,
        maximumFractionDigits: decimals ?? 2,
      })

    default:
      // "Geral" não inventa separador de milhar, mas usa a vírgula decimal. Só
      // número de verdade: "0012" guardado como texto perderia o zero.
      return typeof value === 'number'
        ? value.toLocaleString(LOCALE, { maximumFractionDigits: 10, useGrouping: false })
        : String(value)
  }
}

/** A partir de 1899-12-30: o Excel herdou do Lotus 1-2-3 o 1900 bissexto, que não é. */
const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
const MS_PER_DAY = 86_400_000

export function serialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH + Math.round(serial) * MS_PER_DAY)
}

export function dateToSerial(date: Date): number {
  return Math.round(
    (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - EXCEL_EPOCH) / MS_PER_DAY,
  )
}

function formatDate(serial: number): string {
  if (serial < 1 || serial > 2_958_465) return String(serial)
  return serialToDate(serial).toLocaleDateString(LOCALE, { timeZone: 'UTC' })
}

/** Número, percentual, moeda e data no formato brasileiro; o resto fica texto, nunca número aproximado. */
export function parseInput(raw: string): { value: CellValue; style?: Partial<CellStyle> } {
  const text = raw.trim()
  if (text.length === 0) return { value: '' }

  const percent = /^(-?[\d.,]+)\s*%$/.exec(text)
  if (percent !== null) {
    const number = parseBrazilianNumber(percent[1]!)
    if (number !== null) return { value: number / 100, style: { format: CellFormat.Percent } }
  }

  const currency = /^R\$\s*(-?[\d.,]+)$/i.exec(text)
  if (currency !== null) {
    const number = parseBrazilianNumber(currency[1]!)
    if (number !== null) return { value: number, style: { format: CellFormat.Currency } }
  }

  const date = parseBrazilianDate(text)
  if (date !== null) return { value: dateToSerial(date), style: { format: CellFormat.Date } }

  // Zero à esquerda é intenção: matrícula, CEP, código. Vira texto.
  if (!/^0\d/.test(text)) {
    const number = parseBrazilianNumber(text)
    if (number !== null) return { value: number }
  }

  return { value: text }
}

/**
 * `1.234` é mil duzentos e trinta e quatro, e `1234.56` colado de planilha
 * estrangeira é decimal: ponto seguido de exatamente três dígitos, em todos os
 * grupos, é separador de milhar; o resto é decimal.
 */
export function parseBrazilianNumber(text: string): number | null {
  const trimmed = text.trim()
  if (trimmed.length === 0) return null

  let normalized: string
  if (trimmed.includes(',')) {
    // Com vírgula não há dúvida: vírgula é decimal, ponto é milhar.
    normalized = trimmed.replaceAll('.', '').replace(',', '.')
  } else if (/^-?\d{1,3}(\.\d{3})+$/.test(trimmed)) {
    normalized = trimmed.replaceAll('.', '')
  } else {
    normalized = trimmed
  }

  if (!/^-?\d+(\.\d+)?$/.test(normalized)) return null

  const number = Number(normalized)
  return Number.isFinite(number) ? number : null
}

function parseBrazilianDate(text: string): Date | null {
  const match = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(text)
  if (match === null) return null

  const day = Number(match[1])
  const month = Number(match[2])
  let year = Number(match[3])

  // Ano de dois dígitos: a mesma janela que o Excel usa.
  if (year < 100) year += year < 30 ? 2000 : 1900

  const date = new Date(year, month - 1, day)
  const valid = date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
  return valid ? date : null
}
