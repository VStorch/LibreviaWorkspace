import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { BUNDLED_FONT_FILES } from '@services/document/fonts.js'

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..', '..')

describe('fontes empacotadas', () => {
  it('todo arquivo declarado existe em resources/fonts', () => {
    // O modo de falha sem este teste é silencioso e caro: a regra @font-face
    // aponta para um arquivo que não viajou no instalador, o Chromium substitui
    // por conta própria, e o documento pagina diferente na máquina de quem
    // instala e na de quem programou — que tem as fontes do sistema e nunca vê
    // o problema.
    const faltando = BUNDLED_FONT_FILES.filter(
      (arquivo) => !existsSync(join(raiz, 'resources', 'fonts', arquivo)),
    )

    expect(faltando).toEqual([])
  })
})
