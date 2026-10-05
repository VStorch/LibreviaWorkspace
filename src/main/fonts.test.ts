import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { BUNDLED_FONT_FILES } from '@services/document/fonts.js'

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..', '..')

describe('fontes empacotadas', () => {
  it('todo arquivo declarado existe em resources/fonts', () => {
    // Without this test the failure is silent and costly: the @font-face rule points to a file that
    // did not ship in the installer, Chromium substitutes on its own, and the document paginates
    // differently on the installer's machine than on the developer's, which has the system fonts
    // and never sees it.
    const faltando = BUNDLED_FONT_FILES.filter(
      (arquivo) => !existsSync(join(raiz, 'resources', 'fonts', arquivo)),
    )

    expect(faltando).toEqual([])
  })
})
