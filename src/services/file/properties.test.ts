import { describe, expect, it } from 'vitest'
import {
  CustomPropertyType,
  customPropertiesProblem,
  instantOfDay,
  localDayOf,
  sameCustomProperties,
  stampProperties,
  w3cdtf,
} from './properties.js'

const now = new Date('2026-10-02T12:34:56.789Z')

describe('carimbo das propriedades na gravação', () => {
  it('não carimba o documento que ninguém editou', () => {
    const current = { title: 'T', modified: '2020-01-01T00:00:00Z' }
    expect(stampProperties(current, { author: 'Ana', now, fresh: false, edited: false })).toBe(current)
    expect(stampProperties(undefined, { author: 'Ana', now, fresh: false, edited: false })).toBeUndefined()
  })

  it('o editado ganha quem modificou, quando e a revisão seguinte, e guarda o resto', () => {
    const stamped = stampProperties(
      { title: 'T', creator: 'Bia', created: '2020-01-01T00:00:00Z', revision: '7' },
      { author: ' Ana ', now, fresh: false, edited: true },
    )
    expect(stamped).toEqual({
      title: 'T',
      creator: 'Bia',
      created: '2020-01-01T00:00:00Z',
      modified: '2026-10-02T12:34:56Z',
      lastModifiedBy: 'Ana',
      revision: '8',
    })
  })

  it('o documento novo ganha criador e data de criação', () => {
    expect(stampProperties(undefined, { author: 'Ana', now, fresh: true, edited: false })).toEqual({
      created: '2026-10-02T12:34:56Z',
      creator: 'Ana',
      modified: '2026-10-02T12:34:56Z',
      lastModifiedBy: 'Ana',
      revision: '1',
    })
  })

  it('o arquivo que pede para não guardar nomes não é assinado', () => {
    const stamped = stampProperties(
      { removePersonalInformation: true, creator: '' },
      { author: 'Ana', now, fresh: true, edited: true },
    )
    expect(stamped?.creator).toBe('')
    expect(stamped?.lastModifiedBy).toBeUndefined()
    expect(stamped?.modified).toBe('2026-10-02T12:34:56Z')
  })

  it('sem nome de autor, não assina', () => {
    const stamped = stampProperties({}, { author: '', now, fresh: true, edited: true })
    expect(stamped?.creator).toBeUndefined()
    expect(stamped?.lastModifiedBy).toBeUndefined()
  })

  it('a data sai em W3CDTF sem milissegundos', () => {
    expect(w3cdtf(now)).toBe('2026-10-02T12:34:56Z')
  })
})

describe('propriedades personalizadas', () => {
  const texto = { name: 'Cliente', type: CustomPropertyType.Text, value: 'XPTO' }
  const numero = { name: 'Processo', type: CustomPropertyType.Number, value: '2026' }

  it('aceita a lista válida', () => {
    expect(customPropertiesProblem([texto, numero])).toBeNull()
  })

  it('recusa nome vazio, nome repetido sem olhar maiúsculas, número inválido e data vazia', () => {
    expect(customPropertiesProblem([{ ...texto, name: '  ' }])).toBe('nameMissing')
    expect(customPropertiesProblem([texto, { ...numero, name: 'cliente' }])).toBe('nameRepeated')
    expect(customPropertiesProblem([{ ...numero, value: 'doze' }])).toBe('numberInvalid')
    expect(customPropertiesProblem([{ ...numero, value: '' }])).toBe('numberInvalid')
    expect(customPropertiesProblem([{ name: 'Prazo', type: CustomPropertyType.Date, value: '' }])).toBe(
      'dateMissing',
    )
  })

  it('a mesma lista em outra ordem é a mesma', () => {
    expect(sameCustomProperties([texto, numero], [numero, texto])).toBe(true)
    expect(sameCustomProperties([texto], [{ ...texto, value: 'outro' }])).toBe(false)
    expect(sameCustomProperties([texto], [])).toBe(false)
  })

  it('a data vai e volta como o mesmo dia no fuso de quem usa', () => {
    const instant = instantOfDay('2026-10-12')
    expect(instant).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:00Z$/)
    expect(localDayOf(instant)).toBe('2026-10-12')
    expect(instantOfDay('12/10/2026')).toBe('')
    expect(localDayOf('não é data')).toBe('')
  })
})
