/// <reference lib="dom" />
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles } from './fixtures.js'

/**
 * Referências (M8): painel de navegação, marcadores, sumário, legendas e
 * referências cruzadas.
 *
 * A preferência é mudada por `window.api.preferences.set`, e não pelo menu
 * nativo, pelo mesmo motivo de `exibir.spec.ts`: o Playwright não alcança o menu
 * do sistema, e o caminho depois do clique é o mesmo.
 */
async function setPreference(session: Session, patch: Record<string, unknown>): Promise<void> {
  await session.window.evaluate(async (value) => {
    const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } }).api
    await api.preferences.set(value)
  }, patch)
}

test.describe('painel de navegação', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-referencias-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('lista os títulos ao vivo, leva a eles e destaca o da seção do cursor', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()

    await session.window.keyboard.press('Control+Alt+1')
    await session.window.keyboard.type('Capítulo um')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Texto do capítulo.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Control+Alt+2')
    await session.window.keyboard.type('Seção dois')

    await setPreference(session, { navigationPane: true })
    const pane = session.window.getByRole('navigation', { name: 'Navegação' })
    await expect(pane).toBeVisible()

    const entries = pane.getByRole('button', { name: /^Nível/ })
    await expect(entries).toHaveText(['Capítulo um', 'Seção dois'])
    // O cursor está no fim do segundo título: é a seção dele que se destaca.
    await expect(entries.nth(1)).toHaveAttribute('aria-current', 'location')

    await entries.nth(0).click()
    await expect(entries.nth(0)).toHaveAttribute('aria-current', 'location')
    // O cursor foi para o título: digitar escreve nele.
    await session.window.keyboard.type('>')
    await expect(session.window.locator('.ProseMirror h1')).toHaveText('>Capítulo um')

    await setPreference(session, { navigationPane: false })
    await expect(pane).toHaveCount(0)
  })

  test('no documento travado também leva ao título', async () => {
    // O comentário trava a edição. O título é o `Ttulo1` do Word em português, que
    // o painel reconhece pelo nome interno `heading 1`.
    const origem = join(folder, 'travado.docx')
    const p = (style: string, text: string): string =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    await writeFile(
      origem,
      await docxWithNamedStyles(
        Array.from({ length: 40 }, (_, index) => p('Normal', `Parágrafo ${index + 1}.`)).join('') +
          p('Ttulo1', 'Conclusão') +
          '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Comentado.</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>',
      ),
    )
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.banner--readonly')).toBeVisible()

    await setPreference(session, { navigationPane: true })
    const entries = session.window
      .getByRole('navigation', { name: 'Navegação' })
      .getByRole('button', { name: /^Nível/ })
    await expect(entries).toHaveText(['Relatório anual', 'Conclusão'])

    await entries.nth(1).click()
    await expect(entries.nth(1)).toHaveAttribute('aria-current', 'location')
    // Rolou até o título, que estava folhas abaixo.
    await expect(session.window.locator('.ProseMirror h1', { hasText: 'Conclusão' })).toBeInViewport()
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})
