import { describe, expect, it } from 'vitest'
import { CellFormat } from './model.js'
import { dateToSerial, formatCell, parseBrazilianNumber, parseInput, serialToDate } from './format.js'

describe('formatCell', () => {
  it('mostra célula vazia como texto vazio', () => {
    expect(formatCell(undefined)).toBe('')
    expect(formatCell({})).toBe('')
  })

  it('não inventa separador de milhar no formato geral', () => {
    // O usuário digitou 1000 e espera ver 1000.
    expect(formatCell({ value: 1000 })).toBe('1000')
  })

  it('formata moeda em reais', () => {
    expect(formatCell({ value: 1234.5, style: { format: CellFormat.Currency } })).toMatch(/R\$\s?1\.234,50/)
  })

  it('formata percentual multiplicando por cem', () => {
    // O valor guardado é 0,15 — o "15%" é aparência.
    expect(formatCell({ value: 0.15, style: { format: CellFormat.Percent } })).toBe('15%')
  })

  it('respeita o número de casas decimais', () => {
    expect(formatCell({ value: 3.14159, style: { format: CellFormat.Number, decimals: 2 } })).toBe('3,14')
  })

  it('formata número com separador brasileiro', () => {
    expect(formatCell({ value: 1234.5, style: { format: CellFormat.Number, decimals: 1 } })).toBe('1.234,5')
  })

  it('mostra texto como está, mesmo parecendo número', () => {
    expect(formatCell({ value: '0012', style: { format: CellFormat.Text } })).toBe('0012')
  })

  it('traduz booleano', () => {
    expect(formatCell({ value: true })).toBe('VERDADEIRO')
    expect(formatCell({ value: false })).toBe('FALSO')
  })
})

describe('datas', () => {
  it('converte ida e volta', () => {
    const date = new Date(2026, 7, 15)
    expect(serialToDate(dateToSerial(date)).getUTCDate()).toBe(15)
  })

  it('usa a mesma origem do Excel', () => {
    // 1 de janeiro de 2026 é o número de série 46023 no Excel. A origem é
    // 1899-12-30 por causa do bug de 1900 bissexto, herdado do Lotus 1-2-3.
    expect(dateToSerial(new Date(2026, 0, 1))).toBe(46023)
  })

  it('formata data pelo número de série', () => {
    expect(formatCell({ value: 46023, style: { format: CellFormat.Date } })).toBe('01/01/2026')
  })
})

describe('parseInput', () => {
  it('reconhece número simples', () => {
    expect(parseInput('42').value).toBe(42)
  })

  it('reconhece número no formato brasileiro', () => {
    expect(parseInput('1.234,56').value).toBe(1234.56)
  })

  it('reconhece percentual e guarda a fração', () => {
    const parsed = parseInput('15%')
    expect(parsed.value).toBe(0.15)
    expect(parsed.style?.format).toBe(CellFormat.Percent)
  })

  it('reconhece moeda', () => {
    const parsed = parseInput('R$ 1.500,00')
    expect(parsed.value).toBe(1500)
    expect(parsed.style?.format).toBe(CellFormat.Currency)
  })

  it('reconhece data brasileira', () => {
    const parsed = parseInput('15/08/2026')
    expect(parsed.style?.format).toBe(CellFormat.Date)
    expect(serialToDate(parsed.value as number).getUTCMonth()).toBe(7)
  })

  it('preserva zero à esquerda como texto', () => {
    // Matrícula, CEP e código começam com zero. Virar número seria perda
    // silenciosa de dado — e o usuário só descobriria ao imprimir.
    expect(parseInput('0012').value).toBe('0012')
  })

  it.each(['abc', '12abc', '1,2,3', ''])('deixa %o como texto', (input) => {
    const parsed = parseInput(input)
    expect(typeof parsed.value).toBe('string')
  })

  it('recusa data impossível em vez de deslizar o mês', () => {
    // 31/02 no JavaScript vira 3 de março se não for conferido.
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
  ])('converte %s em %f', (text, expected) => {
    // O ponto é ambíguo e a ambiguidade custa caro: no Brasil "1.234" é mil
    // duzentos e trinta e quatro, mas "1234.56" colado de fora é decimal.
    // Escolher um significado fixo erraria metade dos casos por mil vezes.
    expect(parseBrazilianNumber(text)).toBe(expected)
  })

  it.each(['', 'abc', '1,2,3', '1..2'])('recusa %o', (text) => {
    expect(parseBrazilianNumber(text)).toBeNull()
  })
})
