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

describe('property stamp on save', () => {
  it('does not stamp a document nobody edited', () => {
    const current = { title: 'T', modified: '2020-01-01T00:00:00Z' }
    expect(stampProperties(current, { author: 'Ana', now, fresh: false, edited: false })).toBe(current)
    expect(stampProperties(undefined, { author: 'Ana', now, fresh: false, edited: false })).toBeUndefined()
  })

  it('an edited document gets who modified it, when and the next revision, and keeps the rest', () => {
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

  it('a new document gets a creator and creation date', () => {
    expect(stampProperties(undefined, { author: 'Ana', now, fresh: true, edited: false })).toEqual({
      created: '2026-10-02T12:34:56Z',
      creator: 'Ana',
      modified: '2026-10-02T12:34:56Z',
      lastModifiedBy: 'Ana',
      revision: '1',
    })
  })

  it('a file that asks not to keep names is not signed', () => {
    const stamped = stampProperties(
      { removePersonalInformation: true, creator: '' },
      { author: 'Ana', now, fresh: true, edited: true },
    )
    expect(stamped?.creator).toBe('')
    expect(stamped?.lastModifiedBy).toBeUndefined()
    expect(stamped?.modified).toBe('2026-10-02T12:34:56Z')
  })

  it('without an author name, does not sign', () => {
    const stamped = stampProperties({}, { author: '', now, fresh: true, edited: true })
    expect(stamped?.creator).toBeUndefined()
    expect(stamped?.lastModifiedBy).toBeUndefined()
  })

  it('the date comes out in W3CDTF without milliseconds', () => {
    expect(w3cdtf(now)).toBe('2026-10-02T12:34:56Z')
  })
})

describe('custom properties', () => {
  const texto = { name: 'Cliente', type: CustomPropertyType.Text, value: 'XPTO' }
  const numero = { name: 'Processo', type: CustomPropertyType.Number, value: '2026' }

  it('accepts the valid list', () => {
    expect(customPropertiesProblem([texto, numero])).toBeNull()
  })

  it('rejects an empty name, a duplicate name ignoring case, an invalid number and an empty date', () => {
    expect(customPropertiesProblem([{ ...texto, name: '  ' }])).toBe('nameMissing')
    expect(customPropertiesProblem([texto, { ...numero, name: 'cliente' }])).toBe('nameRepeated')
    expect(customPropertiesProblem([{ ...numero, value: 'doze' }])).toBe('numberInvalid')
    expect(customPropertiesProblem([{ ...numero, value: '' }])).toBe('numberInvalid')
    expect(customPropertiesProblem([{ name: 'Prazo', type: CustomPropertyType.Date, value: '' }])).toBe(
      'dateMissing',
    )
  })

  it('the same list in another order is the same', () => {
    expect(sameCustomProperties([texto, numero], [numero, texto])).toBe(true)
    expect(sameCustomProperties([texto], [{ ...texto, value: 'outro' }])).toBe(false)
    expect(sameCustomProperties([texto], [])).toBe(false)
  })

  it("the date round-trips as the same day in the user's time zone", () => {
    const instant = instantOfDay('2026-10-12')
    expect(instant).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:00Z$/)
    expect(localDayOf(instant)).toBe('2026-10-12')
    expect(instantOfDay('12/10/2026')).toBe('')
    expect(localDayOf('não é data')).toBe('')
  })
})
