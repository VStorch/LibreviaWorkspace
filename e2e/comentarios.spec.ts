import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCommentThread, entryOf } from './fixtures.js'

/**
 * Comentários (M10): lidos, mostrados ao lado da folha e devolvidos ao arquivo
 * quando o parágrafo que os ancora é editado (fase 1); criados, respondidos,
 * resolvidos e excluídos no painel, e gravados de volta (fase 2).
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
  test('criar, responder, resolver e excluir, e o arquivo e o rascunho guardam tudo', async () => {
    const origem = join(pasta, 'ata.docx')
    const rascunho = join(pasta, 'ata.sdoc')
    const destino = join(pasta, 'volta.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, save: rascunho, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    // Novo comentário sobre "Fim.": o cartão abre com a caixa de texto em foco.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    const caixa = session.window.locator('.comment-composer__field')
    await expect(caixa).toBeFocused()
    await session.window.keyboard.type('Revisar o fim.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes).toHaveCount(3)
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(editor.locator('.comment-range').filter({ hasText: 'Fim.' })).toHaveCount(1)

    // Esc desiste do cartão novo, e as pontas saem com ele.
    await editor.getByText('Ata da reunião').click()
    await menu(session, 'insert-comment')
    await expect(caixa).toBeFocused()
    await session.window.keyboard.press('Escape')
    await expect(cartoes).toHaveCount(3)

    // Responder e resolver a conversa da Ana.
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await conversa.click()
    await conversa.getByRole('button', { name: 'Responder' }).click()
    await session.window.keyboard.type('Fechado.')
    await session.window.keyboard.press('Control+Enter')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Fechado.')
    await conversa.getByRole('button', { name: 'Resolver' }).click()
    await expect(conversa).toHaveClass(/comment-card--done/)

    // Excluir a conversa resolvida da Carla.
    const resolvida = cartoes.filter({ hasText: 'Carla' })
    await resolvida.click()
    await resolvida.getByRole('button', { name: 'Excluir' }).click()
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Carla' })).toHaveCount(0)

    // Desfazer a exclusão devolve o cartão; refazer o tira de novo.
    await editor.click()
    await session.window.keyboard.press('Control+Z')
    await expect(cartoes).toHaveCount(3)
    await session.window.keyboard.press('Control+Shift+Z')
    await expect(cartoes).toHaveCount(2)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const comentarios = await entryOf(origem, 'word/comments.xml')
    expect(comentarios).toContain('Revisar o fim.')
    expect(comentarios).toContain('Fechado.')
    expect(comentarios).not.toContain('Trocar o título.')
    expect(await entryOf(origem, 'word/commentsExtended.xml')).toMatch(/w15:paraId="10000000" w15:done="1"/)
    expect(await entryOf(origem, 'word/document.xml')).not.toContain('w:id="2"')

    // O rascunho guarda o mesmo estado, e o .docx gravado dele também.
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: rascunho, save: destino })
    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()
    await menu(session, 'open')
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(cartoes.filter({ hasText: 'Ana' })).toHaveClass(/comment-card--done/)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: destino })
    await menu(session, 'open')
    await expect(cartoes).toHaveCount(2)
    const ana = cartoes.filter({ hasText: 'Ana' })
    await expect(ana).toHaveClass(/comment-card--done/)
    await ana.click()
    await expect(ana.locator('.comment-card__replies')).toContainText('Conferido na planilha.')
    await expect(ana.locator('.comment-card__replies')).toContainText('Fechado.')
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
  })
})
