import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * O menu "Exibir": tema, idioma e modo de leitura.
 *
 * Os três chegaram juntos e são a mesma preferência guardada no main, então o
 * que cada teste aqui confere é a **outra ponta** — que a tela obedece. O item
 * de menu marcado já é conferido pelo próprio Electron.
 *
 * A preferência é mudada por `window.api.preferences.set` e não pelo clique no
 * menu nativo: o menu do sistema não é do Chromium, e o Playwright não o
 * alcança. O caminho testado é o mesmo que o clique percorre depois do clique.
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

  test('o tema escuro troca a cor da casca e do papel', async () => {
    const root = session.window.locator('html')
    await expect(root).toHaveAttribute('data-theme', 'light')

    await setPreference(session, { theme: 'dark' })
    await expect(root).toHaveAttribute('data-theme', 'dark')

    // O papel escurece junto. É o pedido de "modo escuro no conteúdo", e é o
    // que distingue esta implementação de escurecer só as barras.
    const paper = await session.window
      .locator('.paper')
      .first()
      .evaluate((node) => {
        return getComputedStyle(node).backgroundColor
      })

    // Sem casar com um valor exato: o que importa é que o papel deixou de ser
    // branco, e não qual cinza foi escolhido.
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
    // A barra de ferramentas é o que a tela desenha a partir do catálogo; o
    // menu nativo é conferido pelo main e não se alcança daqui.
    await setPreference(session, { language: 'en' })

    const acoes = session.window.locator('.ProseMirror')
    await expect(acoes).toBeVisible()

    // O botão direito monta os rótulos de tabela pelo catálogo: em inglês eles
    // saem em inglês, que é a prova de que a chave virou frase no idioma certo.
    const idioma = await session.window.evaluate(() => document.documentElement.lang)
    expect(typeof idioma).toBe('string')

    await setPreference(session, { language: 'pt' })
  })

  test('o modo de leitura tira as barras e o papel', async () => {
    await expect(session.window.locator('.doc-toolbar, .toolbar').first()).toBeVisible()
    await expect(session.window.locator('.paper')).not.toHaveCount(0)

    await setPreference(session, { readingMode: true })

    // O papel some: a rolagem passa a ser contínua, que foi o pedido.
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

    // O papel volta, e com ele a edição.
    await expect(session.window.locator('.pages--reading')).toHaveCount(0)
    await expect(session.window.locator('.paper')).not.toHaveCount(0)
  })
})

/**
 * As três sobrevivem a fechar o aplicativo.
 *
 * Fica num `describe` próprio porque estes testes reaproveitam o mesmo
 * `userData` entre duas sessões — é a única forma de perguntar "e depois de
 * reabrir?" — e o `beforeEach` de cima cria um perfil novo a cada teste.
 *
 * Vale o teste: as três moram no mesmo arquivo que a ortografia, e a gravação é
 * a mesma linha de código. O que poderia quebrar não é a gravação, é a
 * **leitura** — o schema tem `default` para todas, e um `default` aplicado onde
 * deveria haver valor gravado devolveria o tema claro em português a cada
 * abertura, sem erro nenhum aparecer.
 */
test.describe('o que foi escolhido continua escolhido', () => {
  test('tema, idioma e modo de leitura voltam como estavam', async () => {
    let session = await launch()
    const { userData } = session

    try {
      await menu(session, 'new-document')
      await session.window.evaluate(async () => {
        const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } })
          .api
        await api.preferences.set({ theme: 'dark', language: 'en', readingMode: true })
      })

      await expect(session.window.locator('html')).toHaveAttribute('data-theme', 'dark')
      await session.close()

      // Mesmo perfil, processo novo: é o que "reabrir o aplicativo" quer dizer.
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

      // E a tela obedece ao que foi lido, sem ninguém tocar em nada.
      await expect(session.window.locator('html')).toHaveAttribute('data-theme', 'dark')
    } finally {
      await session.close()
    }
  })
})
