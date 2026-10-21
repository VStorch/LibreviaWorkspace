import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCommentThread, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/**
 * Comments: read, shown beside the sheet and returned to the file when the paragraph anchoring them
 * is edited; created, replied to, resolved and deleted in the pane, and saved back; the cursor
 * around anchors, cut and paste, navigation, hiding the pane, the author and paper.
 */
test.describe('comments', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-comments-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('the thread appears beside the excerpt, and editing the paragraph does not lose it', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()

    // One thread per card: the reply lives inside the question's.
    const cards = session.window.locator('.comment-card')
    await expect(cards).toHaveCount(2)
    const conversation = cards.filter({ hasText: 'Conferir o valor.' })
    await expect(conversation).toContainText('Ana')
    await expect(conversation.locator('.comment-card__replies')).toContainText('Conferido na planilha.')

    // A resolved one comes collapsed: author and notice, without the text.
    const resolved = cards.filter({ hasText: 'Resolvido' })
    await expect(resolved).toContainText('Carla')
    await expect(resolved).not.toContainText('Trocar o título.')

    // At the height of the range it comments.
    const excerpt = editor.locator('.comment-range').first()
    await expect(excerpt).toHaveText('doze mil reais')
    await expect
      .poll(async () => {
        const [excerptBox, cardBox] = await Promise.all([excerpt.boundingBox(), conversation.boundingBox()])
        return Math.abs(excerptBox!.y - cardBox!.y)
      })
      .toBeLessThan(30)

    // Clicking the card selects and highlights the range.
    await conversation.click()
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')

    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect(session.window.locator('.banner--notice')).toBeHidden()

    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).toContain('Revisado.')
    for (const id of ['0', '1', '2']) {
      expect(body).toContain(`<w:commentRangeStart w:id="${id}"`)
      expect(body).toContain(`<w:commentReference w:id="${id}"`)
    }

    // Reopened, the card is still there, with the reply.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(cards).toHaveCount(2)
    await expect(cards.filter({ hasText: 'Conferir o valor.' })).toContainText('Conferido na planilha.')
  })
  test('create, reply, resolve and delete, and the file and the draft keep everything', async () => {
    const sourceFile = join(directory, 'ata.docx')
    const draft = join(directory, 'ata.sdoc')
    const destinationFile = join(directory, 'volta.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, save: draft, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cards = session.window.locator('.comment-card')
    await expect(cards).toHaveCount(2)

    // A new comment on "Fim.": the card opens with the text box focused.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    const box = session.window.locator('.comment-composer__field')
    await expect(box).toBeFocused()
    await session.window.keyboard.type('Revisar o fim.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cards).toHaveCount(3)
    await expect(cards.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(editor.locator('.comment-range').filter({ hasText: 'Fim.' })).toHaveCount(1)

    // Esc gives up the new card, and its ends go with it.
    await editor.getByText('Ata da reunião').click()
    await menu(session, 'insert-comment')
    await expect(box).toBeFocused()
    await session.window.keyboard.press('Escape')
    await expect(cards).toHaveCount(3)

    const conversation = cards.filter({ hasText: 'Conferir o valor.' })
    await conversation.click()
    await conversation.getByRole('button', { name: 'Responder' }).click()
    await session.window.keyboard.type('Fechado.')
    await session.window.keyboard.press('Control+Enter')
    await expect(conversation.locator('.comment-card__replies')).toContainText('Fechado.')
    await conversation.getByRole('button', { name: 'Resolver' }).click()
    await expect(conversation).toHaveClass(/comment-card--done/)

    const resolved = cards.filter({ hasText: 'Carla' })
    await resolved.click()
    await resolved.getByRole('button', { name: 'Excluir' }).click()
    await expect(cards).toHaveCount(2)
    await expect(cards.filter({ hasText: 'Carla' })).toHaveCount(0)

    // Undoing the deletion brings the card back; redoing removes it again.
    await editor.click()
    await session.window.keyboard.press('Control+Z')
    await expect(cards).toHaveCount(3)
    await session.window.keyboard.press('Control+Shift+Z')
    await expect(cards).toHaveCount(2)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const comments = await entryOf(sourceFile, 'word/comments.xml')
    expect(comments).toContain('Revisar o fim.')
    expect(comments).toContain('Fechado.')
    expect(comments).not.toContain('Trocar o título.')
    expect(await entryOf(sourceFile, 'word/commentsExtended.xml')).toMatch(
      /w15:paraId="10000000" w15:done="1"/,
    )
    expect(await entryOf(sourceFile, 'word/document.xml')).not.toContain('w:id="2"')

    // The draft keeps the same state, and so does the .docx saved from it.
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: draft, save: destinationFile })
    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()
    await menu(session, 'open')
    await expect(cards).toHaveCount(2)
    await expect(cards.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(cards.filter({ hasText: 'Ana' })).toHaveClass(/comment-card--done/)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: destinationFile })
    await menu(session, 'open')
    await expect(cards).toHaveCount(2)
    const ana = cards.filter({ hasText: 'Ana' })
    await expect(ana).toHaveClass(/comment-card--done/)
    await ana.click()
    await expect(ana.locator('.comment-card__replies')).toContainText('Conferido na planilha.')
    await expect(ana.locator('.comment-card__replies')).toContainText('Fechado.')
    await expect(cards.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
  })
  test('a point comment at the end of the paragraph does not misalign the selection', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cards = session.window.locator('.comment-card')
    await expect(cards).toHaveCount(2)

    // A point comment at the end of "Fim.": only the `commentEnd`, touching the end.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Ponto.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cards).toHaveCount(3)

    // Selecting only the last character is not the paragraph: deleting it keeps the comment.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+ArrowLeft')
    await session.window.keyboard.press('Delete')
    await expect(editor.locator('p').last()).toHaveText('Fim')
    await expect(cards).toHaveCount(3)
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
    await expect(cards.filter({ hasText: 'Carla' })).toHaveCount(1)
  })
  test('cut and paste carries the comment; copy and paste does not repeat it', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cards = session.window.locator('.comment-card')
    const carla = cards.filter({ hasText: 'Carla' })
    await expect(carla).toHaveCount(1)

    // The whole line carries the ends touching it: once cut, the thread goes with it.
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
    expect(await entryOf(sourceFile, 'word/comments.xml')).toContain('Trocar o título.')
    expect(
      (await entryOf(sourceFile, 'word/document.xml')).match(/<w:commentRangeStart w:id="2"/g),
    ).toHaveLength(1)
  })

  test('next and previous pick the thread in text order', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const active = session.window.locator('.comment-card--active')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    // Bruno's reply has its own ends in the file, but is not a thread.
    await menu(session, 'next-comment')
    await expect(active).toContainText('Ana')
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')
    await menu(session, 'next-comment')
    await expect(active).toContainText('Carla')
    await menu(session, 'next-comment')
    await expect(active).toContainText('Ana')
    await menu(session, 'previous-comment')
    await expect(active).toContainText('Carla')

    // The cursor went along: what is typed lands in the chosen range.
    await expect(editor).toBeFocused()
  })

  test('hiding the panel removes the cards and highlight, and the file keeps everything', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const panel = session.window.locator('.comments-pane')
    await expect(panel).toHaveCount(1)
    await expect(editor.locator('.comment-range')).not.toHaveCount(0)

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(panel).toHaveCount(0)
    await expect(editor.locator('.comment-range')).toHaveCount(0)

    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).toContain('<w:commentRangeStart w:id="0"')
    expect(body).toContain('<w:commentRangeStart w:id="2"')
    expect(await entryOf(sourceFile, 'word/comments.xml')).toContain('Conferir o valor.')

    // The preference sticks: with the document reopened, the pane stays hidden.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(panel).toHaveCount(0)

    // Inserting a comment brings the pane back.
    await editor.getByText('Fim.').click()
    await menu(session, 'insert-comment')
    await expect(panel).toHaveCount(1)
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.press('Escape')

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(panel).toHaveCount(0)
    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(panel).toHaveCount(1)
    await expect(editor.locator('.comment-range').first()).toBeVisible()
  })

  test('the author name signs the new comment, and the reply can be edited', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cards = session.window.locator('.comment-card')
    await expect(cards).toHaveCount(2)

    await menu(session, 'author-name')
    const dialog = session.window.getByRole('dialog', { name: 'Nome do autor' })
    const name = dialog.getByRole('textbox', { name: 'Nome' })
    await expect(name).toBeFocused()
    await name.fill('Zé da Silva')
    await dialog.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialog).toHaveCount(0)

    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Assinado.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cards.filter({ hasText: 'Assinado.' })).toContainText('Zé da Silva')

    // Bruno's reply, edited in Ana's card.
    const conversation = cards.filter({ hasText: 'Conferir o valor.' })
    await conversation.click()
    const replies = conversation.locator('.comment-card__replies')
    await replies.getByRole('button', { name: 'Editar' }).click()
    const box = replies.locator('.comment-composer__field')
    await expect(box).toBeFocused()
    await expect(box).toHaveValue('Conferido na planilha.')
    await box.fill('Conferido duas vezes.')
    await session.window.keyboard.press('Control+Enter')
    await expect(replies).toContainText('Conferido duas vezes.')
    await expect(replies).not.toContainText('Conferido na planilha.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const comments = await entryOf(sourceFile, 'word/comments.xml')
    expect(comments).toMatch(
      /w:author="Zé da Silva"[^>]*w:initials="ZDS"|w:initials="ZDS"[^>]*w:author="Zé da Silva"/,
    )
    expect(comments).toContain('Conferido duas vezes.')
    expect(comments).not.toContain('Conferido na planilha.')
  })

  test('the PDF comes out without the comments', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    const sourceFile = join(directory, 'ata.docx')
    const destinationFile = join(directory, 'ata.pdf')
    await writeFile(sourceFile, await docxWithCommentThread())
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    await menu(session, 'export-pdf')
    await expect.poll(() => pdfText(destinationFile), { timeout: 30_000 }).toContain('doze mil reais')
    const text = await pdfText(destinationFile)
    for (const outside of [
      'Conferir o valor.',
      'Conferido na planilha.',
      'Ana',
      'Bruno',
      'Carla',
      'Resolvido',
    ])
      expect(text).not.toContain(outside)
  })
})
