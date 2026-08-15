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

describe('ida e volta do formato interno', () => {
  it('preserva a planilha inteira', () => {
    expect(parseWorkbook(serializeWorkbook(rich))).toEqual(rich)
  })

  it('preserva larguras de coluna com a chave numérica', () => {
    // As chaves saem do JSON como texto. Sem a conversão de volta,
    // `columnWidths[0]` procuraria um número e encontraria a chave "0".
    const restored = parseWorkbook(serializeWorkbook(rich))

    expect(restored.sheets[0]!.columnWidths[0]).toBe(180)
    expect(restored.sheets[0]!.columnWidths[3]).toBe(60)
  })

  it('preserva congelamento e formatação de célula', () => {
    const sheet = parseWorkbook(serializeWorkbook(rich)).sheets[0]!

    expect(sheet.frozenRows).toBe(1)
    expect(sheet.frozenColumns).toBe(1)
    expect(sheet.cells['C2']?.style?.format).toBe(CellFormat.Currency)
    expect(sheet.cells['A1']?.style?.bold).toBe(true)
  })

  it('preserva uma planilha vazia', () => {
    const empty = createEmptyWorkbook()
    expect(parseWorkbook(serializeWorkbook(empty))).toEqual(empty)
  })

  it('grava o mapa esparso, sem linhas vazias', () => {
    // Uma planilha de mil linhas com duas células preenchidas não pode gerar
    // mil entradas no arquivo.
    const raw = JSON.parse(serializeWorkbook(rich)) as { sheets: { cells: object }[] }

    expect(Object.keys(raw.sheets[0]!.cells)).toEqual(['A1', 'C2'])
  })

  it('grava a versão do formato', () => {
    expect(JSON.parse(serializeWorkbook(createEmptyWorkbook()))).toMatchObject({
      format: 'ssheet',
      version: SSHEET_VERSION,
    })
  })
})

describe('leitura de arquivo problemático', () => {
  it('recusa JSON malformado com mensagem compreensível', () => {
    expect(() => parseWorkbook('{ isto não é json')).toThrow(AppError)
    expect(() => parseWorkbook('{ isto não é json')).toThrow(/corrompido|válida/i)
  })

  it('recusa JSON válido que não é planilha', () => {
    expect(() => parseWorkbook('{"qualquer":"coisa"}')).toThrow(/planilha válida/i)
  })

  it('recusa versão futura em vez de adivinhar', () => {
    const future = JSON.stringify({
      format: 'ssheet',
      version: SSHEET_VERSION + 1,
      sheets: [createSheet('A')],
      activeSheet: 0,
    })

    expect(() => parseWorkbook(future)).toThrow(/versão mais recente/i)
  })

  it('recupera aba ativa fora do intervalo sem descartar os dados', () => {
    // O dado vale mais que a lembrança de qual aba estava aberta.
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

  it('abre planilha gravada sem os campos acrescentados depois', () => {
    // Compatibilidade: acrescentar campo opcional não pode invalidar o que já
    // está em disco.
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
  it('reconhece o formato pelo conteúdo', () => {
    expect(isSpreadsheetFile(serializeWorkbook(createEmptyWorkbook()))).toBe(true)
  })

  it.each(['', 'texto solto', '{"format":"sdoc"}'])('recusa %o', (text) => {
    expect(isSpreadsheetFile(text)).toBe(false)
  })
})
