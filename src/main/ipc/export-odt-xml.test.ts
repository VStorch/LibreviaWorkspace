import { execFileSync } from 'node:child_process'
import { access, constants, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, type DocumentModel } from '@services/document/model.js'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { sdocContentOf } from '../docx/index.js'
import { exportOdt } from '../odt/index.js'
import { SidecarClient } from '../sidecar/client.js'
import { sidecarPathIn } from '../sidecar/locate.js'

/** The system tools, when there are: `xmllint` checks, `unzip` takes the parts out. */
const hasTool = (tool: string, flag: string): boolean => {
  try {
    execFileSync(tool, [flag], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}

const HAS_TOOLS = hasTool('xmllint', '--version') && hasTool('unzip', '-v')

const executable = sidecarPathIn(process.cwd())
const published = await access(executable, constants.X_OK).then(
  () => true,
  () => false,
)

const client = new SidecarClient(() => Promise.resolve(executable))
afterAll(() => client.dispose())

/** Here, not next to the writer, because only main has the published sidecar and `xmllint`. */
describe.skipIf(!published || !HAS_TOOLS)('the ODT with equations', () => {
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
          attrs: { sectionBreak: 's1' },
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

  it('every XML part is well formed', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'librevia-odt-xml-'))
    try {
      const path = join(directory, 'equações.odt')
      await writeFile(path, await exportOdt(client, sdocContentOf(model), 'Librevia/0.0.0'))
      const parts = execFileSync('unzip', ['-Z1', path], { encoding: 'utf8' })
        .split('\n')
        .filter((name) => name.endsWith('.xml'))
      expect(parts).toContain('Object 2/content.xml')
      for (const part of parts) {
        const data = execFileSync('unzip', ['-p', path, part])
        expect(
          () => execFileSync('xmllint', ['--noout', '-'], { input: data, stdio: ['pipe', 'ignore', 'pipe'] }),
          part,
        ).not.toThrow()
      }
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })
})
