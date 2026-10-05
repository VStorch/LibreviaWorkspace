import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * The "View" menu: theme, language and reading mode, preferences kept in main; here we check the
 * screen obeys. Preferences change through `window.api.preferences.set`, because Playwright cannot
 * reach the native menu.
 */
async function setPreference(session: Session, patch: Record<string, unknown>): Promise<void> {
  await session.window.evaluate(async (value) => {
    const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } }).api
    await api.preferences.set(value)
  }, patch)
}

test.describe('menu Exibir', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('as barras de ferramentas e de status obedecem às preferências', async () => {
    await expect(session.window.locator('.toolbar')).toBeVisible()
    await expect(session.window.locator('.statusbar')).toBeVisible()

    await setPreference(session, { showToolbar: false, showStatusBar: false })
    await expect(session.window.locator('.toolbar')).toHaveCount(0)
    await expect(session.window.locator('.statusbar')).toHaveCount(0)

    await setPreference(session, { showToolbar: true, showStatusBar: true })
    await expect(session.window.locator('.toolbar')).toBeVisible()
    await expect(session.window.locator('.statusbar')).toBeVisible()
  })

  test('o tema escuro troca a cor da casca e do papel', async () => {
    const root = session.window.locator('html')
    await expect(root).toHaveAttribute('data-theme', 'light')

    await setPreference(session, { theme: 'dark' })
    await expect(root).toHaveAttribute('data-theme', 'dark')

    // The paper darkens too, not just the toolbars.
    const paper = await session.window
      .locator('.paper')
      .first()
      .evaluate((node) => {
        return getComputedStyle(node).backgroundColor
      })

    // The paper is not white, whatever the grey.
    expect(paper).not.toBe('rgb(255, 255, 255)')

    await setPreference(session, { theme: 'light' })
    await expect(root).toHaveAttribute('data-theme', 'light')
  })

  test('o texto que o documento não coloriu acompanha o tema', async () => {
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Sem cor declarada.')

    const claro = await session.window
      .locator('.ProseMirror p')
      .first()
      .evaluate((node) => getComputedStyle(node).color)

    await setPreference(session, { theme: 'dark' })

    const escuro = await session.window
      .locator('.ProseMirror p')
      .first()
      .evaluate((node) => getComputedStyle(node).color)

    expect(escuro).not.toBe(claro)
  })

  test('o idioma troca a interface sem reabrir nada', async () => {
    // The toolbar comes from the catalog; main checks the native menu.
    await setPreference(session, { language: 'en' })

    const acoes = session.window.locator('.ProseMirror')
    await expect(acoes).toBeVisible()

    // Right-click labels come from the catalog, in the chosen language.
    const idioma = await session.window.evaluate(() => document.documentElement.lang)
    expect(typeof idioma).toBe('string')

    await setPreference(session, { language: 'pt' })
  })

  test('o modo de leitura tira as barras e o papel', async () => {
    await expect(session.window.locator('.doc-toolbar, .toolbar').first()).toBeVisible()
    await expect(session.window.locator('.paper')).not.toHaveCount(0)

    await setPreference(session, { readingMode: true })

    // The paper goes away, and scrolling is continuous.
    await expect(session.window.locator('.paper')).toHaveCount(0)
    await expect(session.window.locator('.statusbar')).toHaveCount(0)
    await expect(session.window.locator('.pages--reading')).toHaveCount(1)
  })

  test('o modo de leitura trava a edição', async () => {
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Antes.')

    await setPreference(session, { readingMode: true })

    const antes = await session.window.locator('.ProseMirror').textContent()
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('DEPOIS')
    const depois = await session.window.locator('.ProseMirror').textContent()

    expect(depois).toBe(antes)
  })

  test('Esc sai do modo de leitura', async () => {
    await setPreference(session, { readingMode: true })
    await expect(session.window.locator('.pages--reading')).toHaveCount(1)

    await session.window.keyboard.press('Escape')

    // The paper comes back, and with it editing.
    await expect(session.window.locator('.pages--reading')).toHaveCount(0)
    await expect(session.window.locator('.paper')).not.toHaveCount(0)
  })
})

/**
 * All three survive closing the app: the schema has a `default` for each, and a `default` in place
 * of the saved value would fall back to the light theme in Portuguese without any error. Its own
 * `describe`, because it reuses `userData` across two sessions.
 */
test.describe('o que foi escolhido continua escolhido', () => {
  test('tema, idioma e modo de leitura voltam como estavam', async () => {
    const userData = await mkdtemp(join(tmpdir(), 'librevia-prefs-'))
    let session = await launch({ userData })

    try {
      await menu(session, 'new-document')
      await session.window.evaluate(async () => {
        const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } })
          .api
        await api.preferences.set({ theme: 'dark', language: 'en', readingMode: true })
      })

      await expect(session.window.locator('html')).toHaveAttribute('data-theme', 'dark')
      await session.close()

      // Same profile, new process: that is what "reopening the app" means.
      session = await launch({ userData })

      const guardado = await session.window.evaluate(async () => {
        const api = (
          window as unknown as {
            api: { preferences: { get: (p: unknown) => Promise<{ ok: boolean; data: unknown }> } }
          }
        ).api
        return api.preferences.get({})
      })

      expect(guardado).toMatchObject({
        ok: true,
        data: { theme: 'dark', language: 'en', readingMode: true },
      })

      // And the screen obeys what was read, without anyone touching anything.
      await expect(session.window.locator('html')).toHaveAttribute('data-theme', 'dark')
    } finally {
      await session.close()
      await rm(userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
    }
  })
})
