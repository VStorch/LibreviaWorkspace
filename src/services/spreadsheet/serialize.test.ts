import { describe, expect, it } from 'vitest'
import { AppError } from '@shared/errors.js'
import { CellFormat, createEmptyWorkbook, createSheet, setCell, type WorkbookModel } from './model.js'
import { SSHEET_VERSION, isSpreadsheetFile, parseWorkbook, serializeWorkbook } from './serialize.js'

const rich: WorkbookModel = {
  sheets: [
    {
      ...setCell(setCell(createSheet('Vendas'), 0, 0, { value: 'Produto', style: { bold: true } }), 1, 2, {
        value: 1500,
        style: { format: CellFormat.Currency, decimals: 2, align: 'right' },
      }),
      columnWidths: { 0: 180, 3: 60 },
      frozenRows: 1,
      frozenColumns: 1,
    },
    createSheet('Resumo'),
  ],
  activeSheet: 1,
}

describe('round trip of the internal format', () => {
  it('preserves the whole sheet', () => {
    expect(parseWorkbook(serializeWorkbook(rich))).toEqual(rich)
  })

  it('preserves column widths with the numeric key', () => {
    // Keys come out of JSON as text. Without converting back, `columnWidths[0]` would look for a
    // number and find the key "0".
    const restored = parseWorkbook(serializeWorkbook(rich))

    expect(restored.sheets[0]!.columnWidths[0]).toBe(180)
    expect(restored.sheets[0]!.columnWidths[3]).toBe(60)
  })

  it('preserves freeze panes and cell formatting', () => {
    const sheet = parseWorkbook(serializeWorkbook(rich)).sheets[0]!

    expect(sheet.frozenRows).toBe(1)
    expect(sheet.frozenColumns).toBe(1)
    expect(sheet.cells['C2']?.style?.format).toBe(CellFormat.Currency)
    expect(sheet.cells['A1']?.style?.bold).toBe(true)
  })

  it('preserves an empty sheet', () => {
    const empty = createEmptyWorkbook()
    expect(parseWorkbook(serializeWorkbook(empty))).toEqual(empty)
  })

  it('writes the sparse map, without empty rows', () => {
    // A thousand-row sheet with two filled cells must not produce a thousand entries in the file.
    const raw = JSON.parse(serializeWorkbook(rich)) as { sheets: { cells: object }[] }

    expect(Object.keys(raw.sheets[0]!.cells)).toEqual(['A1', 'C2'])
  })

  it('writes the format version', () => {
    expect(JSON.parse(serializeWorkbook(createEmptyWorkbook()))).toMatchObject({
      format: 'ssheet',
      version: SSHEET_VERSION,
    })
  })
})

describe('reading a problematic file', () => {
  it('rejects malformed JSON with an understandable message', () => {
    expect(() => parseWorkbook('{ isto não é json')).toThrow(AppError)
    expect(() => parseWorkbook('{ isto não é json')).toThrow(/corrompido|válida/i)
  })

  it('rejects valid JSON that is not a spreadsheet', () => {
    expect(() => parseWorkbook('{"qualquer":"coisa"}')).toThrow(/planilha válida/i)
  })

  it('rejects a future version instead of guessing', () => {
    const future = JSON.stringify({
      format: 'ssheet',
      version: SSHEET_VERSION + 1,
      sheets: [createSheet('A')],
      activeSheet: 0,
    })

    expect(() => parseWorkbook(future)).toThrow(/versão mais recente/i)
  })

  it('recovers an out-of-range active sheet without discarding the data', () => {
    // The data is worth more than remembering which tab was open.
    const broken = JSON.stringify({
      format: 'ssheet',
      version: SSHEET_VERSION,
      sheets: [createSheet('Única')],
      activeSheet: 7,
    })

    const restored = parseWorkbook(broken)
    expect(restored.activeSheet).toBe(0)
    expect(restored.sheets).toHaveLength(1)
  })

  it('opens a sheet saved without the fields added later', () => {
    // Compatibility: adding an optional field must not invalidate what is already on disk.
    const older = JSON.stringify({
      format: 'ssheet',
      version: SSHEET_VERSION,
      sheets: [{ name: 'A', cells: { A1: { value: 1 } } }],
    })

    const sheet = parseWorkbook(older).sheets[0]!
    expect(sheet.frozenRows).toBe(0)
    expect(sheet.rowCount).toBeGreaterThan(0)
  })
})

describe('isSpreadsheetFile', () => {
  it('recognizes the format by its content', () => {
    expect(isSpreadsheetFile(serializeWorkbook(createEmptyWorkbook()))).toBe(true)
  })

  it.each(['', 'texto solto', '{"format":"sdoc"}'])('rejects %o', (text) => {
    expect(isSpreadsheetFile(text)).toBe(false)
  })
})

describe('sheet properties', () => {
  it('come back from the file as they were written', () => {
    const properties = { title: 'Orçamento 2026', creator: 'Ana', revision: '3' }
    const reopened = parseWorkbook(serializeWorkbook({ ...createEmptyWorkbook(), properties }))

    expect(reopened.properties).toEqual(properties)
  })

  it('a version 1 sheet opens without properties, and the XLSX keeps the ones it had', () => {
    const version1 = JSON.stringify({ format: 'ssheet', version: 1, sheets: [createSheet('A')] })

    expect(parseWorkbook(version1)).not.toHaveProperty('properties')
  })
})
