import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCommentThread, entryOf } from './fixtures.js'

/**
 * Comentários (M10, fase 1): lidos, mostrados ao lado da folha e devolvidos ao
 * arquivo quando o parágrafo que os ancora é editado.
 */
test.describe('comentários', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-comentarios-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('a conversa aparece ao lado do trecho, e editar o parágrafo não a perde', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()

    // Uma conversa por cartão: a resposta mora dentro do da pergunta.
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await expect(conversa).toContainText('Ana')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Conferido na planilha.')

    // A resolvida vem recolhida: o autor e o aviso, sem o texto.
    const resolvida = cartoes.filter({ hasText: 'Resolvido' })
    await expect(resolvida).toContainText('Carla')
    await expect(resolvida).not.toContainText('Trocar o título.')

    // Na altura do trecho que comenta.
    const trecho = editor.locator('.comment-range').first()
    await expect(trecho).toHaveText('doze mil reais')
    await expect
      .poll(async () => {
        const [caixaTrecho, caixaCartao] = await Promise.all([trecho.boundingBox(), conversa.boundingBox()])
        return Math.abs(caixaTrecho!.y - caixaCartao!.y)
      })
      .toBeLessThan(30)

    // Clicar no cartão seleciona o trecho e o realça.
    await conversa.click()
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')

    // Editar o parágrafo comentado e gravar.
    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect(session.window.locator('.banner--notice')).toBeHidden()

    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('Revisado.')
    for (const id of ['0', '1', '2']) {
      expect(corpo).toContain(`<w:commentRangeStart w:id="${id}"`)
      expect(corpo).toContain(`<w:commentReference w:id="${id}"`)
    }

    // Reaberto, o cartão continua lá, com a resposta.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Conferir o valor.' })).toContainText('Conferido na planilha.')
  })
})
