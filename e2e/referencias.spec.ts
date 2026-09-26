/// <reference lib="dom" />
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, docxWithReferences, entryOf } from './fixtures.js'

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

/**
 * Marcadores sem sumário nem campos: os dois ainda travavam o documento quando
 * este teste nasceu, e aqui o que se quer é editar.
 */
const BOOKMARKS_BODY =
  '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc100"/><w:r><w:t>Introdução</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
  '<w:p><w:bookmarkStart w:id="1" w:name="Resumo"/><w:r><w:t xml:space="preserve">O resumo começa aqui </w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>e termina aqui.</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>' +
  '<w:p><w:r><w:t xml:space="preserve">Veja o </w:t></w:r><w:hyperlink w:anchor="Resumo" w:history="1"><w:r><w:t>resumo</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> e a conclusão.</w:t></w:r></w:p>'

test.describe('marcadores', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-marcadores-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('lista, vai para, adiciona e grava sem perder os ocultos', async () => {
    const origem = join(folder, 'marcadores.docx')
    const destino = join(folder, 'saida.docx')
    await writeFile(origem, await docxWithReferences(BOOKMARKS_BODY))
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('O resumo começa aqui')
    await expect(editor).toHaveAttribute('contenteditable', 'true')

    await menu(session, 'insert-bookmark')
    const dialog = session.window.getByRole('dialog', { name: 'Marcadores' })
    const list = dialog.getByRole('listbox', { name: 'Marcadores do documento' })
    // O oculto do sumário só aparece com a caixa marcada.
    await expect(list.getByRole('option')).toHaveText(['Resumo'])
    await dialog.getByLabel('Marcadores ocultos').check()
    await expect(list.getByRole('option')).toHaveText(['_Toc100', 'Resumo'])

    await list.getByRole('option', { name: 'Resumo' }).click()
    await dialog.getByRole('button', { name: 'Ir para' }).click()
    await expect(dialog).toHaveCount(0)
    await session.window.keyboard.type('>')
    await expect(editor.locator('p', { hasText: 'começa aqui' })).toHaveText('>O resumo começa aqui ')

    // Um marcador novo sobre o ponto final do último parágrafo.
    const last = editor.locator('p', { hasText: 'e a conclusão' })
    await last.dblclick({ position: { x: 5, y: 5 } })
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+ArrowLeft')
    await menu(session, 'insert-bookmark')
    await dialog.getByLabel('Nome do marcador').fill('com espaço')
    await expect(dialog.getByRole('button', { name: 'Adicionar' })).toBeDisabled()
    await dialog.getByLabel('Nome do marcador').fill('Conclusao')
    await dialog.getByRole('button', { name: 'Adicionar' }).click()
    await expect(dialog).toHaveCount(0)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toMatch(/w:name="Conclusao"/)
    // O oculto do título e a ponta que mora no corpo, entre dois parágrafos.
    expect(corpo).toMatch(/w:name="_Toc100"/)
    expect(corpo).toMatch(/<\/w:p><w:bookmarkEnd w:id="1" ?\/>/)
    expect(corpo).toMatch(/w:anchor="Resumo"/)
  })

  test('Ctrl+clique no link interno leva ao marcador', async () => {
    const origem = join(folder, 'link.docx')
    await writeFile(origem, await docxWithReferences(BOOKMARKS_BODY))
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')

    await editor.locator('a', { hasText: 'resumo' }).click({ modifiers: ['Control'] })
    await session.window.keyboard.type('>')
    await expect(editor.locator('p', { hasText: 'começa aqui' })).toHaveText('>O resumo começa aqui ')
  })

  test('link para um título cria o marcador oculto em volta dele', async () => {
    const origem = join(folder, 'titulo.docx')
    const destino = join(folder, 'titulo-saida.docx')
    await writeFile(
      origem,
      await docxWithReferences(BOOKMARKS_BODY.replace(/<w:bookmark(Start|End)[^>]*\/>/g, '')),
    )
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')

    // Cursor no fim do último parágrafo, sem seleção: o link leva o texto do título.
    await editor.locator('p', { hasText: 'e a conclusão' }).click()
    await session.window.keyboard.press('End')
    await session.window.getByRole('button', { name: 'Link' }).click()
    const dialog = session.window.getByRole('dialog', { name: 'Inserir link' })
    await dialog.getByLabel('Lugar neste documento').selectOption({ label: 'Introdução' })
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('a', { hasText: 'Introdução' })).toHaveAttribute('href', /^#_Ref\d+$/)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    const nome = /w:anchor="(_Ref\d+)"/.exec(corpo)?.[1]
    expect(nome).toBeDefined()
    expect(corpo).toContain(`w:name="${nome ?? ''}"`)
  })
})
