import { describe, expect, it } from 'vitest'
import type { DocumentProperties } from '@services/file/properties.js'
import { propertyFieldChoices, propertyFieldResult } from './property-fields.js'

const properties: DocumentProperties = {
  title: 'contrato de locação',
  creator: 'Ana',
  description: 'Rascunho',
  custom: [
    { name: 'Cliente', type: 'text', value: 'XPTO' },
    { name: 'Valor', type: 'number', value: '12.5' },
    { name: 'Prazo', type: 'date', value: '2026-10-12T15:00:00Z' },
    { name: 'Aprovado', type: 'boolean', value: 'true' },
  ],
}

const result = (instr: string): string | null => propertyFieldResult(instr, properties, 'pt')

describe('propertyFieldResult', () => {
  it('os campos próprios e o DOCPROPERTY leem as mesmas propriedades', () => {
    expect(result(' TITLE ')).toBe('contrato de locação')
    expect(result(' AUTHOR \\* MERGEFORMAT ')).toBe('Ana')
    expect(result(' DOCPROPERTY Comments ')).toBe('Rascunho')
    expect(result(' DOCPROPERTY Company ')).toBe('')
  })

  it('acha a personalizada sem olhar maiúsculas e formata número, data e sim ou não', () => {
    expect(result(' DOCPROPERTY "cliente" ')).toBe('XPTO')
    expect(result(' DOCPROPERTY Valor ')).toBe('12,5')
    expect(result(' DOCPROPERTY Prazo ')).toBe(new Date('2026-10-12T15:00:00Z').toLocaleDateString('pt'))
    expect(result(' DOCPROPERTY Aprovado ')).toBe('Y')
  })

  it('propriedade que não existe dá null', () => {
    expect(result(' DOCPROPERTY "Processo" ')).toBeNull()
  })

  it('segue as opções de maiúsculas do Word', () => {
    expect(result(' TITLE \\* Upper \\* MERGEFORMAT ')).toBe('CONTRATO DE LOCAÇÃO')
    expect(result(' TITLE \\* Caps ')).toBe('Contrato De Locação')
    expect(result(' TITLE \\* FirstCap ')).toBe('Contrato de locação')
  })
})

describe('propertyFieldChoices', () => {
  it('oferece as do resumo e as personalizadas, sem nome com aspas', () => {
    const choices = propertyFieldChoices({
      custom: [...properties.custom!, { name: 'a"b', type: 'text', value: '' }],
    })
    expect(choices[0]).toEqual({ builtIn: true, key: 'title', instr: ' TITLE \\* MERGEFORMAT ' })
    expect(choices.filter((choice) => !choice.builtIn).map((choice) => choice.instr)).toEqual([
      ' DOCPROPERTY "Cliente" \\* MERGEFORMAT ',
      ' DOCPROPERTY "Valor" \\* MERGEFORMAT ',
      ' DOCPROPERTY "Prazo" \\* MERGEFORMAT ',
      ' DOCPROPERTY "Aprovado" \\* MERGEFORMAT ',
    ])
  })
})
