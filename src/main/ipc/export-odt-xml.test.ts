import { execFileSync } from 'node:child_process'
import { describe, expect, it } from 'vitest'
import { odtEntries } from '@services/document/export-odt.js'
import { DEFAULT_PAGE_SETUP, type DocumentModel } from '@services/document/model.js'
import { BUILTIN_STYLES } from '@services/document/styles.js'

/** The system `xmllint`, when there is one. */
const HAS_XMLLINT = (() => {
  try {
    execFileSync('xmllint', ['--version'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
})()

/** Here, not next to the export, because only main has `xmllint`. */
describe('o ODT com equações', () => {
  const model: DocumentModel = {
    page: DEFAULT_PAGE_SETUP,
    sections: [{ ...DEFAULT_PAGE_SETUP, id: 's1' }],
    styles: BUILTIN_STYLES,
    properties: {},
    doc: {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Seja <x> & ' },
            {
              type: 'math',
              attrs: {
                mathml: '<math><msup><mi>x</mi><mn>2</mn></msup></math>',
                latex: 'x<2 & y',
                display: false,
              },
            },
          ],
        },
        {
          type: 'paragraph',
          content: [
            {
              type: 'math',
              attrs: {
                mathml: '<math display="block"><mfrac><mi>a</mi><mo>&lt;</mo></mfrac></math>',
                display: true,
              },
            },
          ],
        },
      ],
    },
  }

  it.skipIf(!HAS_XMLLINT)('toda parte de XML é bem formada', () => {
    const parts = odtEntries(model).filter((entry) => entry.name.endsWith('.xml'))
    expect(parts.map((entry) => entry.name)).toContain('Object 2/content.xml')
    for (const entry of parts) {
      expect(
        () =>
          execFileSync('xmllint', ['--noout', '-'], { input: entry.data, stdio: ['pipe', 'ignore', 'pipe'] }),
        entry.name,
      ).not.toThrow()
    }
  })
})
