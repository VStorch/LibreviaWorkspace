import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCommentThread, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/** Clicks a native menu item, by menu and item label. */
async function clickMenuItem(session: Session, menuLabel: string, itemLabel: string): Promise<void> {
  await session.app.evaluate(
    ({ Menu }, [top, item]) => {
      const submenu = Menu.getApplicationMenu()?.items.find((entry) => entry.label === top)?.submenu
      submenu?.items.find((entry) => entry.label === item)?.click()
    },
    [menuLabel, itemLabel] as const,
  )
}

/**
 * Comments: read, shown beside the sheet and returned to the file when the paragraph anchoring them
 * is edited; created, replied to, resolved and deleted in the pane, and saved back; the cursor
 * around anchors, cut and paste, navigation, hiding the pane, the author and paper.
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

    // One thread per card: the reply lives inside the question's.
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await expect(conversa).toContainText('Ana')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Conferido na planilha.')

    // A resolved one comes collapsed: author and notice, without the text.
    const resolvida = cartoes.filter({ hasText: 'Resolvido' })
    await expect(resolvida).toContainText('Carla')
    await expect(resolvida).not.toContainText('Trocar o título.')

    // At the height of the range it comments.
    const trecho = editor.locator('.comment-range').first()
    await expect(trecho).toHaveText('doze mil reais')
    await expect
      .poll(async () => {
        const [caixaTrecho, caixaCartao] = await Promise.all([trecho.boundingBox(), conversa.boundingBox()])
        return Math.abs(caixaTrecho!.y - caixaCartao!.y)
      })
      .toBeLessThan(30)

    // Clicking the card selects and highlights the range.
    await conversa.click()
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')

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

    // Reopened, the card is still there, with the reply.
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

    // A new comment on "Fim.": the card opens with the text box focused.
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

    // Esc gives up the new card, and its ends go with it.
    await editor.getByText('Ata da reunião').click()
    await menu(session, 'insert-comment')
    await expect(caixa).toBeFocused()
    await session.window.keyboard.press('Escape')
    await expect(cartoes).toHaveCount(3)

    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await conversa.click()
    await conversa.getByRole('button', { name: 'Responder' }).click()
    await session.window.keyboard.type('Fechado.')
    await session.window.keyboard.press('Control+Enter')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Fechado.')
    await conversa.getByRole('button', { name: 'Resolver' }).click()
    await expect(conversa).toHaveClass(/comment-card--done/)

    const resolvida = cartoes.filter({ hasText: 'Carla' })
    await resolvida.click()
    await resolvida.getByRole('button', { name: 'Excluir' }).click()
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Carla' })).toHaveCount(0)

    // Undoing the deletion brings the card back; redoing removes it again.
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

    // The draft keeps the same state, and so does the .docx saved from it.
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
  test('o comentário de ponto no fim do parágrafo não desencontra a seleção', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    // A point comment at the end of "Fim.": only the `commentEnd`, touching the end.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Ponto.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes).toHaveCount(3)

    // Selecting only the last character is not the paragraph: deleting it keeps the comment.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+ArrowLeft')
    await session.window.keyboard.press('Delete')
    await expect(editor.locator('p').last()).toHaveText('Fim')
    await expect(cartoes).toHaveCount(3)
    await session.window.keyboard.type('.')

    // `End` and `Shift+Home` in quick succession select the line, and a single `Backspace` deletes
    // it.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type('Novo')
    await expect(editor.locator('p').last()).toHaveText('Novo')

    // Without a selection, `Backspace` skips over the anchor (the end of Carla's comment) and
    // deletes the character, and the thread stays.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Backspace')
    await expect(editor.getByText('Título provisóri', { exact: true })).toBeVisible()
    await expect(cartoes.filter({ hasText: 'Carla' })).toHaveCount(1)
  })
  test('recortar e colar leva o comentário; copiar e colar não o repete', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    const carla = cartoes.filter({ hasText: 'Carla' })
    await expect(carla).toHaveCount(1)

    // A linha inteira leva as pontas encostadas nela: recortada, a conversa sai.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    await session.window.keyboard.press('Control+X')
    await expect(carla).toHaveCount(0)

    // Pasted at the end, it comes back with the card.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Control+V')
    await expect(carla).toHaveCount(1)
    await expect(editor.locator('[data-comment-start][data-cid="2"]')).toHaveCount(1)

    // Copied and pasted again, the text repeats and the thread does not.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    await session.window.keyboard.press('Control+C')
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Control+V')
    await expect(editor.getByText('Título provisório')).toHaveCount(2)
    await expect(editor.locator('[data-comment-start][data-cid="2"]')).toHaveCount(1)
    await expect(carla).toHaveCount(1)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(origem, 'word/comments.xml')).toContain('Trocar o título.')
    expect((await entryOf(origem, 'word/document.xml')).match(/<w:commentRangeStart w:id="2"/g)).toHaveLength(
      1,
    )
  })

  test('próximo e anterior escolhem a conversa pela ordem do texto', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const ativo = session.window.locator('.comment-card--active')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    // Bruno's reply has its own ends in the file, but is not a thread.
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Ana')
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Carla')
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Ana')
    await menu(session, 'previous-comment')
    await expect(ativo).toContainText('Carla')

    // The cursor went along: what is typed lands in the chosen range.
    await expect(editor).toBeFocused()
  })

  test('esconder o painel tira os cartões e o realce, e o arquivo guarda tudo', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const painel = session.window.locator('.comments-pane')
    await expect(painel).toHaveCount(1)
    await expect(editor.locator('.comment-range')).not.toHaveCount(0)

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(0)
    await expect(editor.locator('.comment-range')).toHaveCount(0)

    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('<w:commentRangeStart w:id="0"')
    expect(corpo).toContain('<w:commentRangeStart w:id="2"')
    expect(await entryOf(origem, 'word/comments.xml')).toContain('Conferir o valor.')

    // The preference sticks: with the document reopened, the pane stays hidden.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(painel).toHaveCount(0)

    // Inserting a comment brings the pane back.
    await editor.getByText('Fim.').click()
    await menu(session, 'insert-comment')
    await expect(painel).toHaveCount(1)
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.press('Escape')

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(0)
    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(1)
    await expect(editor.locator('.comment-range').first()).toBeVisible()
  })

  test('o nome do autor assina o comentário novo, e a resposta se edita', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    await menu(session, 'author-name')
    const dialogo = session.window.getByRole('dialog', { name: 'Nome do autor' })
    const nome = dialogo.getByRole('textbox', { name: 'Nome' })
    await expect(nome).toBeFocused()
    await nome.fill('Zé da Silva')
    await dialogo.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialogo).toHaveCount(0)

    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Assinado.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes.filter({ hasText: 'Assinado.' })).toContainText('Zé da Silva')

    // Bruno's reply, edited in Ana's card.
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await conversa.click()
    const respostas = conversa.locator('.comment-card__replies')
    await respostas.getByRole('button', { name: 'Editar' }).click()
    const caixa = respostas.locator('.comment-composer__field')
    await expect(caixa).toBeFocused()
    await expect(caixa).toHaveValue('Conferido na planilha.')
    await caixa.fill('Conferido duas vezes.')
    await session.window.keyboard.press('Control+Enter')
    await expect(respostas).toContainText('Conferido duas vezes.')
    await expect(respostas).not.toContainText('Conferido na planilha.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const comentarios = await entryOf(origem, 'word/comments.xml')
    expect(comentarios).toMatch(
      /w:author="Zé da Silva"[^>]*w:initials="ZDS"|w:initials="ZDS"[^>]*w:author="Zé da Silva"/,
    )
    expect(comentarios).toContain('Conferido duas vezes.')
    expect(comentarios).not.toContain('Conferido na planilha.')
  })

  test('o PDF sai sem os comentários', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext não instalado')
    const origem = join(pasta, 'ata.docx')
    const destino = join(pasta, 'ata.pdf')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    await menu(session, 'export-pdf')
    await expect.poll(() => pdfText(destino), { timeout: 30_000 }).toContain('doze mil reais')
    const texto = await pdfText(destino)
    for (const fora of ['Conferir o valor.', 'Conferido na planilha.', 'Ana', 'Bruno', 'Carla', 'Resolvido'])
      expect(texto).not.toContain(fora)
  })
})
