import { describe, expect, it } from 'vitest'
import { CellFormat } from './model.js'
import { dateToSerial, formatCell, parseBrazilianNumber, parseInput, serialToDate } from './format.js'

describe('formatCell', () => {
  it('shows an empty cell as empty text', () => {
    expect(formatCell(undefined)).toBe('')
    expect(formatCell({})).toBe('')
  })

  it('invents no thousands separator in the general format', () => {
    // The user typed 1000 and expects to see 1000.
    expect(formatCell({ value: 1000 })).toBe('1000')
  })

  it('uses a decimal comma in the general format', () => {
    // The result of an average almost always lands here: showing 147.43 in a Brazilian app is the
    // same kind of mistake as showing the date backwards.
    expect(formatCell({ value: 884.6 })).toBe('884,6')
    expect(formatCell({ value: 147.43333333333334 })).toBe('147,4333333333')
  })

  it('but does not convert text that looks like a number, even in general', () => {
    // An ID "0012" would lose the leading zero if converted.
    expect(formatCell({ value: '0012' })).toBe('0012')
  })

  it('formats currency in reais', () => {
    expect(formatCell({ value: 1234.5, style: { format: CellFormat.Currency } })).toMatch(/R\$\s?1\.234,50/)
  })

  it('formats a percentage by multiplying by one hundred', () => {
    // The stored value is 0.15; "15%" is appearance.
    expect(formatCell({ value: 0.15, style: { format: CellFormat.Percent } })).toBe('15%')
  })

  it('respects the number of decimal places', () => {
    expect(formatCell({ value: 3.14159, style: { format: CellFormat.Number, decimals: 2 } })).toBe('3,14')
  })

  it('formats a number with the Brazilian separator', () => {
    expect(formatCell({ value: 1234.5, style: { format: CellFormat.Number, decimals: 1 } })).toBe('1.234,5')
  })

  it('shows text as is, even if it looks like a number', () => {
    expect(formatCell({ value: '0012', style: { format: CellFormat.Text } })).toBe('0012')
  })

  it('translates a boolean', () => {
    expect(formatCell({ value: true })).toBe('VERDADEIRO')
    expect(formatCell({ value: false })).toBe('FALSO')
  })
})

describe('dates', () => {
  it('converts back and forth', () => {
    const date = new Date(2026, 7, 15)
    expect(serialToDate(dateToSerial(date)).getUTCDate()).toBe(15)
  })

  it('uses the same epoch as Excel', () => {
    // January 1, 2026 is serial number 46023 in Excel. The origin is 1899-12-30 because of the 1900
    // leap year bug inherited from Lotus 1-2-3.
    expect(dateToSerial(new Date(2026, 0, 1))).toBe(46023)
  })

  it('formats a date from its serial number', () => {
    expect(formatCell({ value: 46023, style: { format: CellFormat.Date } })).toBe('01/01/2026')
  })
})

describe('parseInput', () => {
  it('recognizes a plain number', () => {
    expect(parseInput('42').value).toBe(42)
  })

  it('recognizes a number in Brazilian format', () => {
    expect(parseInput('1.234,56').value).toBe(1234.56)
  })

  it('recognizes a percentage and stores the fraction', () => {
    const parsed = parseInput('15%')
    expect(parsed.value).toBe(0.15)
    expect(parsed.style?.format).toBe(CellFormat.Percent)
  })

  it('recognizes currency', () => {
    const parsed = parseInput('R$ 1.500,00')
    expect(parsed.value).toBe(1500)
    expect(parsed.style?.format).toBe(CellFormat.Currency)
  })

  it('recognizes a Brazilian date', () => {
    const parsed = parseInput('15/08/2026')
    expect(parsed.style?.format).toBe(CellFormat.Date)
    expect(serialToDate(parsed.value as number).getUTCMonth()).toBe(7)
  })

  it('preserves a leading zero as text', () => {
    // IDs, postal codes and codes start with zero. Turning them into numbers would be silent data
    // loss, and the user would only find out when printing.
    expect(parseInput('0012').value).toBe('0012')
  })

  it.each(['abc', '12abc', '1,2,3', ''])('leaves %o as text', (input) => {
    const parsed = parseInput(input)
    expect(typeof parsed.value).toBe('string')
  })

  it('rejects an impossible date instead of rolling over the month', () => {
    // In JavaScript 31/02 becomes March 3 unless checked.
    expect(parseInput('31/02/2026').value).toBe('31/02/2026')
  })
})

describe('parseBrazilianNumber', () => {
  it.each([
    ['1234', 1234],
    ['1.234', 1234],
    ['1.234,56', 1234.56],
    ['-9,5', -9.5],
    ['1234.56', 1234.56],
    ['1.234.567', 1234567],
    ['1.2', 1.2],
  ])('converts %s to %f', (text, expected) => {
    // The dot is ambiguous and the ambiguity is costly: in Brazil "1.234" is one thousand two
    // hundred thirty-four, but "1234.56" pasted from elsewhere is a decimal. A fixed meaning would
    // get half the cases wrong by a factor of a thousand.
    expect(parseBrazilianNumber(text)).toBe(expected)
  })

  it.each(['', 'abc', '1,2,3', '1..2'])('rejects %o', (text) => {
    expect(parseBrazilianNumber(text)).toBeNull()
  })
})
