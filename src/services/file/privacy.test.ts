import { describe, expect, it } from 'vitest'
import {
  PrivacyCategory,
  anonymizeAuthors,
  authorsIn,
  cleanOwnProperties,
  foundCategories,
  inspectOwnFile,
  inspectPdf,
  withoutPersonalNames,
} from './privacy.js'

const ANONYMOUS = 'Autor'

const model = {
  doc: {
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        attrs: { markRevision: { kind: 'ins', author: 'Bruno', date: '2026-01-01T00:00:00Z' } },
        content: [{ type: 'text', text: 'oi', marks: [{ type: 'insertion', attrs: { author: 'Carla' } }] }],
      },
    ],
  },
  comments: [{ id: '1', author: 'Ana', initials: 'AL', date: '', paragraphs: ['ver'], done: false }],
}

describe('inspetor de dados pessoais nos arquivos do próprio editor', () => {
  it('acha os autores de comentários, alterações e marcas de parágrafo', () => {
    expect(authorsIn(model).sort()).toEqual(['Ana', 'Bruno', 'Carla'])
  })

  it('junta os nomes das propriedades sem repetir nem listar o nome genérico', () => {
    const report = inspectOwnFile(
      { creator: 'Ana; Dora', lastModifiedBy: 'Ana', company: 'ACME', title: 'Relatório' },
      { comments: [{ author: ANONYMOUS }, { author: 'Bruno' }] },
      ANONYMOUS,
    )
    expect(report.names).toEqual(['Ana', 'Dora', 'ACME', 'Bruno'])
    expect(foundCategories(report)).toEqual([PrivacyCategory.Names])
  })

  it('sem nomes, não acha nada', () => {
    const report = inspectOwnFile({ title: 'Relatório' }, null, ANONYMOUS)
    expect(foundCategories(report)).toEqual([])
    expect(report.removePersonalInformation).toBe(false)
  })

  it('troca cada autor e as iniciais pelo nome genérico e não mexe no resto', () => {
    const clean = anonymizeAuthors(model, ANONYMOUS)
    expect(authorsIn(clean)).toEqual([ANONYMOUS, ANONYMOUS, ANONYMOUS])
    expect(clean.comments[0]?.initials).toBe('A')
    expect(clean.doc.content[0]?.content[0]?.text).toBe('oi')
    expect(clean.doc.content[0]?.attrs.markRevision.date).toBe('2026-01-01T00:00:00Z')
    expect(authorsIn(model)).toContain('Bruno')
  })

  it('limpa a autoria das propriedades e liga a marca, guardando o título', () => {
    expect(withoutPersonalNames({ title: 'T', creator: 'Ana', company: 'ACME' })).toEqual({
      title: 'T',
      creator: '',
      lastModifiedBy: '',
      company: '',
      manager: '',
      removePersonalInformation: true,
    })
  })

  it('no PDF, mostra o autor e quantas versões anteriores o arquivo guarda', () => {
    const report = inspectPdf({ creator: 'Ana Lima', title: 'Relatório' }, 2, ANONYMOUS)
    expect(report.names).toEqual(['Ana Lima'])
    expect(report.previousVersions).toBe(2)
  })

  it('acha e remove as propriedades personalizadas de um arquivo do próprio editor', () => {
    const properties = { title: 'T', custom: [{ name: 'Cliente', type: 'text' as const, value: 'XPTO' }] }
    expect(inspectOwnFile(properties, null, ANONYMOUS).customProperties).toEqual(['Cliente'])
    expect(cleanOwnProperties(properties, [PrivacyCategory.CustomProperties])).toEqual({
      title: 'T',
      custom: [],
    })
    expect(cleanOwnProperties(properties, [PrivacyCategory.Names]).custom).toHaveLength(1)
  })

  it('lista as categorias na ordem do inspetor', () => {
    expect(
      foundCategories({
        names: [],
        paths: ['C:\\Users\\ana'],
        customProperties: ['Cliente'],
        thumbnail: true,
        imagesWithMetadata: 2,
        removePersonalInformation: false,
      }),
    ).toEqual([
      PrivacyCategory.Paths,
      PrivacyCategory.CustomProperties,
      PrivacyCategory.Thumbnail,
      PrivacyCategory.ImageMetadata,
    ])
  })
})
