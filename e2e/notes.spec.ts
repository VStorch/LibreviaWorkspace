import { execFile } from 'node:child_process'
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote, docxWithLongFootnote, docxWithManyFootnotes, entryOf } from './fixtures.js'
import { hasPdfinfo, hasSoffice, pdfInfo } from './external-tools.js'

/**
 * Footnotes and endnotes: read, numbered and preserved.
 *
 * The reference is a node with the note body inside. Editing the paragraph carrying it does not
 * lose it, which is why a document with notes does not open locked.
 */
test.describe('footnotes', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-notes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('editing the reference paragraph, saving and reopening keeps the note', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithFootnote())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    // The note shows at the foot of the first sheet, with its number.
    const note = session.window.locator('.paper-notes .note-body')
    await expect(note).toContainText('Fonte: ata anterior.')
    await expect(note.locator('.note-number')).toHaveText('1')

    await editor.getByText('Segundo parágrafo').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Revisto: ')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The paragraph was rewritten with the reference; the note, which nobody touched, stays in the
    // file, with the separators.
    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).toContain('Revisto: Segundo parágrafo')
    expect(body).toMatch(/<w:footnoteReference w:id="1" ?\/>/)
    const notes = await entryOf(sourceFile, 'word/footnotes.xml')
    expect(notes).toContain('Fonte: ata anterior.')
    expect(notes).toContain('w:separator')

    await menu(session, 'open')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    await expect(session.window.locator('.paper-notes .note-body')).toContainText('Fonte: ata anterior.')
    await expect(session.window.locator('.banner--notice')).toBeHidden()
  })
})

/**
 * Footnotes and endnotes: the body is edited at the foot of the page, and notes take room on the
 * sheet, so screen and paper break at the same place.
 */
test.describe('notes at the foot of the page', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-notes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const text = (session: Session) => session.window.locator('.pages__column .ProseMirror')
  const notes = (session: Session) => session.window.locator('.paper-notes .note-body')

  test('insert a note by command, write in it, undo and redo', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Ata da reunião')
    await menu(session, 'insert-footnote')

    await expect(text(session).locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    await expect(notes(session)).toHaveCount(1)
    // The cursor is already in the note body.
    await expect(notes(session)).toBeFocused()
    await session.window.keyboard.type('Fonte: livro de atas.')
    await expect(notes(session)).toContainText('Fonte: livro de atas.')
    await expect(notes(session).locator('.note-number')).toHaveText('1')

    // A single history: undoing inside the note undoes in the document.
    await session.window.keyboard.press('Control+z')
    await expect(notes(session)).not.toContainText('Fonte: livro de atas.')
    await session.window.keyboard.press('Control+y')
    await expect(notes(session)).toContainText('Fonte: livro de atas.')

    // Clicking the number goes back to the reference; the next one, at the end, is 2.
    await notes(session).locator('.note-number').click()
    await expect(text(session)).toBeFocused()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' e mais')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Segunda nota.')
    await expect(notes(session)).toHaveCount(2)
    await expect(notes(session).nth(1)).toContainText('Segunda nota.')
    await expect(notes(session).nth(1).locator('.note-number')).toHaveText('2')
  })

  test('deleting the reference deletes the note; saving to .docx and reopening keeps the text', async () => {
    const destinationFile = join(directory, 'ata.docx')
    await stubDialogs(session.app, { save: destinationFile, open: destinationFile, messageBox: 1 })
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Primeira')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota que fica.')
    await notes(session).locator('.note-number').click()
    await session.window.keyboard.type(' segunda')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota que sai.')
    await expect(notes(session)).toHaveCount(2)

    // Backspace right after the reference deletes the reference and the note.
    await notes(session).nth(1).locator('.note-number').click()
    await session.window.keyboard.press('Backspace')
    await expect(text(session).locator('sup.note-ref')).toHaveCount(1)
    await expect(notes(session)).toHaveCount(1)
    await expect(notes(session)).toContainText('Nota que fica.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(destinationFile, 'word/footnotes.xml')
    expect(body).toContain('Nota que fica.')
    expect(body).not.toContain('Nota que sai.')

    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(notes(session)).toHaveCount(1)
    await expect(notes(session)).toContainText('Nota que fica.')
  })

  test('notes push lines, and the screen and the PDF have the same sheets', async () => {
    test.skip(!(await hasPdfinfo()), 'pdfinfo not installed')
    const sourceFile = join(directory, 'longo.docx')
    const pdf = join(directory, 'longo.pdf')
    await writeFile(sourceFile, await docxWithManyFootnotes())
    await stubDialogs(session.app, { open: sourceFile, save: pdf, messageBox: 1 })
    await menu(session, 'open')
    await expect(text(session)).toContainText('24. Os conselheiros')

    const sheets = session.window.locator('.paper')
    await expect(notes(session)).toHaveCount(24)
    const onScreen = await sheets.count()
    expect(onScreen).toBeGreaterThan(2)
    // Every sheet with a reference has the notes area, at the foot.
    await expect(session.window.locator('.paper-notes--footnote')).toHaveCount(onScreen)

    await menu(session, 'export-pdf')
    await expect.poll(() => pdfPages(pdf), { timeout: 30_000 }).toBe(onScreen)
    // Note 1 comes out on paper, on the first sheet.
    const { stdout } = await promisify(execFile)('pdftotext', ['-f', '1', '-l', '1', pdf, '-'])
    expect(stdout).toContain('Nota 1.')

    // And LibreOffice, with the same file, reaches the same sheet count.
    if (await hasSoffice()) {
      const copy = join(directory, 'lo.docx')
      await copyFile(sourceFile, copy)
      await promisify(execFile)(
        'soffice',
        ['--headless', '--convert-to', 'pdf', '--outdir', directory, copy],
        {
          timeout: 120_000,
        },
      )
      expect(await pdfPages(join(directory, 'lo.pdf'))).toBe(onScreen)
    }
  })

  test('with track changes on, what is typed in the note is marked', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Original.')
    await expect(notes(session)).toContainText('Original.')

    await menu(session, 'toggle-track-changes')
    await notes(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Novo')
    await expect(notes(session).locator('ins')).toHaveText(' Novo')
    // A tracked Backspace leaves the deleted text in place, struck through.
    await session.window.keyboard.press('End')
    for (let time = 0; time < ' Novo'.length; time++) await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.press('Backspace')
    await expect(notes(session).locator('del')).toHaveText('.')
    await expect(notes(session)).toContainText('Original. Novo')
  })

  test('the endnote goes after the last paragraph', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Corpo do texto')
    await menu(session, 'insert-endnote')
    await session.window.keyboard.type('Ao fim do documento.')
    await expect(text(session).locator('sup.note-ref')).toHaveAttribute('data-note-number', 'i')

    const area = session.window.locator('.paper-notes--endnote')
    await expect(area).toHaveCount(1)
    await expect(area.locator('.note-body')).toContainText('Ao fim do documento.')
    await expect(area.locator('.note-number')).toHaveText('i')
    // Right below the text, not at the foot of the sheet.
    const paragraph = await text(session).locator('p').first().boundingBox()
    const box = await area.boundingBox()
    expect(box!.y - (paragraph!.y + paragraph!.height)).toBeLessThan(60)

    // Search finds the note text.
    await menu(session, 'find-replace')
    await session.window.keyboard.type('fim do documento')
    await expect(area.locator('.search-hit')).toHaveCount(1)
  })

  test('with several paragraphs, the endnote stays below the last one, also on reopening', async () => {
    const destinationFile = join(directory, 'fim.docx')
    await stubDialogs(session.app, { save: destinationFile, open: destinationFile, messageBox: 1 })
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Primeiro parágrafo')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo parágrafo')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Terceiro')
    await menu(session, 'insert-endnote')
    await session.window.keyboard.type('Fim novo.')

    const belowText = async (): Promise<void> => {
      const area = session.window.locator('.paper-notes--endnote')
      await expect(area).toHaveCount(1)
      const last = await text(session).locator('p').last().boundingBox()
      const box = await area.boundingBox()
      expect(box!.y).toBeGreaterThanOrEqual(last!.y + last!.height - 1)
    }
    await belowText()

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(text(session)).toContainText('Terceiro')
    await belowText()
  })

  test('a note longer than the sheet continues on the next one, without covering the text', async () => {
    const sourceFile = join(directory, 'nota-longa.docx')
    const pdf = join(directory, 'nota-longa.pdf')
    await writeFile(sourceFile, await docxWithLongFootnote())
    await stubDialogs(session.app, { open: sourceFile, save: pdf, messageBox: 1 })
    await menu(session, 'open')
    await expect(text(session)).toContainText('Delta')
    await expect(session.window.locator('.paper-notes__slot--continued')).toHaveCount(1)

    const geometry = await session.window.evaluate(() => {
      const box = (element: Element): { top: number; bottom: number } => {
        const { top, bottom } = element.getBoundingClientRect()
        return { top, bottom }
      }
      const sheets = Array.from(document.querySelectorAll('.paper'), box)
      const blocks = Array.from(document.querySelectorAll('.pages__column .ProseMirror > *'), box)
      const cutouts = Array.from(document.querySelectorAll('.paper-notes__slot'), (slot) => {
        const area = box(slot)
        const visible = Array.from(slot.querySelectorAll('p')).filter((p) => {
          const line = p.getBoundingClientRect()
          return line.bottom > area.top + 1 && line.top < area.bottom - 1
        })
        return {
          ...area,
          first: visible[0]?.textContent ?? '',
          last: visible.at(-1)?.textContent ?? '',
        }
      })
      return { sheets, blocks, cutouts }
    })
    // Each slice within a sheet, without covering any text paragraph.
    for (const cutout of geometry.cutouts) {
      expect(geometry.sheets.some((sheet) => cutout.top >= sheet.top && cutout.bottom <= sheet.bottom)).toBe(
        true,
      )
      for (const block of geometry.blocks) {
        expect(cutout.bottom <= block.top + 1 || cutout.top >= block.bottom - 1).toBe(true)
      }
    }
    // The continuation starts where the first sheet stopped and goes to the end of the note.
    const [first, continuation] = geometry.cutouts
    const lineNumber = (line: string): number => Number(/Linha (\d+)/.exec(line)?.[1] ?? -1)
    expect(lineNumber(continuation!.first)).toBe(lineNumber(first!.last) + 1)
    expect(continuation!.last).toContain('Linha 69')

    if (await hasPdfinfo()) {
      await menu(session, 'export-pdf')
      await expect.poll(() => pdfPages(pdf), { timeout: 30_000 }).toBeGreaterThan(1)
      const { stdout } = await promisify(execFile)('pdftotext', ['-f', '2', '-l', '2', pdf, '-'])
      expect(stdout).toContain(`Linha ${lineNumber(continuation!.first)} `)
      expect(stdout).toContain('Linha 69')
      expect(stdout).not.toContain('Nota longa.')
    }
  })

  test('typing fast in the note past the sheet does not escape into the text', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Corpo')
    await menu(session, 'insert-footnote')
    for (let line = 0; line < 70; line++) {
      await session.window.keyboard.type(`L${line} nota longa que precisa continuar.`)
      await session.window.keyboard.press('Enter')
    }
    await session.window.keyboard.type('Fim da nota.')
    await expect(session.window.locator('.paper-notes__slot--continued').first()).toBeAttached()
    // The live body; continuations are copies of it.
    const note = session.window.locator('.paper-notes__slot:not(.paper-notes__slot--continued) > .note-body')
    await expect(note).toBeFocused()
    await expect(text(session).locator('p').first()).toHaveText('Corpo')
    await expect(text(session).locator('p')).toHaveCount(1)
    await expect(note.locator('p')).toHaveCount(71)
    await expect(notes(session).locator('p').last()).toHaveText('Fim da nota.')
  })

  test('undoing the note with the cursor in it returns focus to the text, and redo works', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota')
    await expect(notes(session)).toBeFocused()
    // The note text, then the note.
    await session.window.keyboard.press('Control+z')
    await session.window.keyboard.press('Control+z')
    await expect(notes(session)).toHaveCount(0)
    await expect(text(session)).toBeFocused()

    await session.window.keyboard.press('Control+y')
    await expect(notes(session)).toHaveCount(1)
    await expect(text(session).locator('sup.note-ref')).toHaveCount(1)
  })

  test('accepting from the context menu the change at the cursor inside the note', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Original.')
    await expect(notes(session)).toContainText('Original.')

    await menu(session, 'toggle-track-changes')
    await notes(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Novo')
    await session.window.keyboard.type(' Outro')
    await expect(notes(session).locator('ins')).toHaveText(' Novo Outro')

    await notes(session).locator('ins').click({ button: 'right' })
    const context = session.window.getByRole('menu', { name: 'Ações do documento' })
    await context.getByRole('menuitem', { name: 'Aceitar alteração' }).click()
    await expect(notes(session).locator('ins')).toHaveCount(0)
    await expect(notes(session)).toContainText('Original. Novo Outro')
  })

  test('a new comment inside the note is refused, and the text stays as it was', async () => {
    // LibreOffice does not open a .docx with `w:commentReference` in `footnotes.xml`.
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Ata')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Fonte da ata.')
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('ArrowLeft')
    for (let time = 0; time < 'ata'.length; time++) await session.window.keyboard.press('Shift+ArrowLeft')
    await menu(session, 'insert-comment')

    await expect(session.window.getByText(/não abriria o arquivo/)).toBeVisible()
    await expect(session.window.locator('.comment-card')).toHaveCount(0)
    await expect(notes(session).locator('.comment-range')).toHaveCount(0)
    await expect(notes(session)).toContainText('Fonte da ata.')
  })

  test('a cross-reference to a note shows the number and F9 updates it', async () => {
    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Primeira')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota alvo.')
    await notes(session).locator('.note-number').click()
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Ver nota ')

    await menu(session, 'insert-cross-reference')
    const ref = session.window.getByRole('dialog', { name: 'Referência cruzada' })
    await ref.getByLabel('Tipo').selectOption('note:footnote')
    await ref.getByLabel('Para qual').selectOption({ label: '1 Nota alvo.' })
    await expect(ref.getByLabel('Inserir referência a')).toHaveValue('number')
    await ref.getByRole('button', { name: 'Inserir' }).click()
    const paragraph = text(session).locator('p', { hasText: 'Ver nota' })
    await expect(paragraph).toHaveText('Ver nota 1')

    // A new note before the cited one: F9 now cites 2.
    await text(session).locator('p').first().click()
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota nova.')
    await expect(text(session).locator('sup.note-ref').nth(1)).toHaveAttribute('data-note-number', '2')
    await menu(session, 'update-fields')
    await expect(paragraph).toHaveText('Ver nota 2')
  })

  test('commenting a word in a paragraph that has a note anchors the comment on it', async () => {
    const sourceFile = join(directory, 'comentario.docx')
    const destinationFile = join(directory, 'comentario-salvo.docx')
    await writeFile(sourceFile, await docxWithManyFootnotes(2))
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(notes(session)).toHaveCount(2)
    // Editing an earlier note, and going back to the text.
    await notes(session).first().click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Editada.')
    await text(session).click()
    await session.window.keyboard.press('Control+Home')
    for (let time = 0; time < '1. '.length; time++) await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Shift+ArrowRight')
    await session.window.keyboard.press('Shift+ArrowRight')
    await menu(session, 'insert-comment')
    await session.window.keyboard.type('Ver isto')
    await session.window.keyboard.press('Control+Enter')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const documentName = await entryOf(destinationFile, 'word/document.xml')
    const paragraph = /<w:p>(?:(?!<w:p>).)*?<\/w:p>/s.exec(documentName)?.[0] ?? ''
    expect(paragraph).toMatch(
      /<w:commentRangeStart[^>]*\/>(<w:r>(?:(?!<\/w:r>).)*<\/w:r>)*?<w:commentRangeEnd/s,
    )
    const commented = /<w:commentRangeStart[^>]*\/>(.*?)<w:commentRangeEnd/s.exec(paragraph)?.[1] ?? ''
    expect(commented.replace(/<[^>]+>/g, '')).toBe('Os')
    expect(paragraph).toContain('<w:commentReference')
  })
})

async function pdfPages(filePath: string): Promise<number> {
  return Number(/Pages:\s+(\d+)/.exec(await pdfInfo(filePath))?.[1] ?? 0)
}
