import { describe, expect, it } from 'vitest'
import { stampProperties, w3cdtf } from './properties.js'

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

  it('sem nome de autor, não assina', () => {
    const stamped = stampProperties({}, { author: '', now, fresh: true, edited: true })
    expect(stamped?.creator).toBeUndefined()
    expect(stamped?.lastModifiedBy).toBeUndefined()
  })

  it('a data sai em W3CDTF sem milissegundos', () => {
    expect(w3cdtf(now)).toBe('2026-10-02T12:34:56Z')
  })
})
