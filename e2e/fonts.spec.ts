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
test.describe('bundled fonts', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('the scheme serves the font the installer ships', async () => {
    const result = await session.window.evaluate(async () => {
      const face = new FontFace(
        'ProvaDeCarregamento',
        "url('librevia-font://fonts/Carlito-Regular.ttf') format('truetype')",
      )
      const loaded = await face.load()
      return loaded.status
    })

    expect(result).toBe('loaded')
  })

  test("the five families are available by the document's name", async () => {
    // `check` only reports faces already loaded, and a family no text uses yet is not.
    const available = await session.window.evaluate(async () => {
      const families = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']
      await Promise.all(families.map((family) => document.fonts.load(`12pt '${family}'`)))
      return families.filter((family) => document.fonts.check(`12pt '${family}'`))
    })

    expect(available).toEqual(['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'])
  })

  test('the scheme does not serve files from outside the fonts folder', async () => {
    // The request comes from CSS, and CSS may come from a document anyone wrote. A normalized `../`
    // is the classic way out of a folder believed closed.
    const leaked = await session.window.evaluate(async () => {
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

    expect(leaked).toBe(false)
  })
})
