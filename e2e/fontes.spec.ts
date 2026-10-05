/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope, which knows neither
// `document` nor `FontFace`. The reference only applies to this file: putting DOM in Node's
// `tsconfig` would make main think it has a window.

import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * Bundled fonts reach the screen: Calibri, Cambria, Arial and Times New Roman do not exist on a
 * clean Linux, and the substitutes ship in the installer (`src/main/fonts.ts`). Loaded by URL, not
 * by name, because `local()` would hide the bundled file on the developer's machine.
 */
test.describe('fontes empacotadas', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('o esquema serve a fonte que o instalador leva', async () => {
    const resultado = await session.window.evaluate(async () => {
      const face = new FontFace(
        'ProvaDeCarregamento',
        "url('librevia-font://fonts/Carlito-Regular.ttf') format('truetype')",
      )
      const carregada = await face.load()
      return carregada.status
    })

    expect(resultado).toBe('loaded')
  })

  test('as cinco famílias ficam disponíveis pelo nome do documento', async () => {
    const disponiveis = await session.window.evaluate(() =>
      ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'].filter((familia) =>
        document.fonts.check(`12pt '${familia}'`),
      ),
    )

    expect(disponiveis).toEqual(['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'])
  })

  test('o esquema não serve arquivo de fora da pasta de fontes', async () => {
    // The request comes from CSS, and CSS may come from a document anyone wrote. A normalized `../`
    // is the classic way out of a folder believed closed.
    const vazou = await session.window.evaluate(async () => {
      try {
        const face = new FontFace(
          'Travessia',
          "url('librevia-font://fonts/../../package.json') format('truetype')",
        )
        await face.load()
        return true
      } catch {
        return false
      }
    })

    expect(vazou).toBe(false)
  })
})
