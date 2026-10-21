import { execFile } from 'node:child_process'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTrackedChange, entryOf } from './fixtures.js'
import { hasPdftotext, hasSoffice, pdfText } from './external-tools.js'

/**
 * Track changes: the revision shows on screen, accepting and rejecting change the text, what is
 * typed with tracking on becomes a revision, and the saved file says the same, in Word and in
 * LibreOffice.
 */
test.describe('revisions', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-revisions-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('accept all, save and reopen: no revision is left', async () => {
    const sourceFile = join(directory, 'revisado.docx')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor.locator('ins.revision').first()).toBeVisible()
    await expect(editor.locator('[data-revision]')).toHaveCount(2)

    await menu(session, 'accept-all-changes')
    await expect(editor.locator('ins.revision, del.revision, [data-revision]')).toHaveCount(0)
    await expect(editor).not.toContainText('Trecho excluído.')
    await expect(editor).not.toContainText('Linha excluída')
    await expect(editor).toContainText('com uma inserção revisada.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).not.toMatch(/<w:(ins|del|moveFrom|moveTo)[ >]/)
    expect(body).toContain('com uma inserção revisada.')

    await menu(session, 'open')
    await expect(editor).toContainText('com uma inserção revisada.')
    await expect(editor.locator('ins.revision, del.revision, [data-revision]')).toHaveCount(0)
  })

  test('rejecting a deletion restores the text, and undo restores it again', async () => {
    const sourceFile = join(directory, 'revisado.docx')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await editor.locator('del.revision').first().click()
    await menu(session, 'reject-change')
    await expect(editor.locator('del.revision', { hasText: 'Trecho excluído.' })).toHaveCount(0)
    await expect(editor).toContainText('Trecho excluído.')

    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('del.revision', { hasText: 'Trecho excluído.' })).toHaveCount(1)
    await session.window.keyboard.press('Control+Shift+z')
    await expect(editor.locator('del.revision', { hasText: 'Trecho excluído.' })).toHaveCount(0)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).not.toContain('<w:delText xml:space="preserve"> Trecho excluído.')
    expect(body).toContain('Trecho excluído.')
    // The same paragraph's insertion is still a revision, with the id it had.
    expect(body).toMatch(/<w:ins [^>]*w:id="1"/)
  })

  test('LibreOffice reads the revisions the saved file carries', async () => {
    test.skip(!(await hasSoffice()), 'no LibreOffice on this machine')
    const sourceFile = join(directory, 'revisado.docx')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    // Editing the revised paragraph: it is rewritten, and revisions come from the marks.
    const editor = session.window.locator('.ProseMirror')
    await editor.getByText('Segundo parágrafo,').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Novo: ')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const copy = join(directory, 'copia.docx')
    await copyFile(sourceFile, copy)
    await promisify(execFile)('soffice', ['--headless', '--convert-to', 'odt', '--outdir', directory, copy], {
      timeout: 120_000,
    })
    const contents = await entryOf(join(directory, 'copia.odt'), 'content.xml')
    // Insertion, deletion, paragraph mark, move (both sides) and row.
    const regions = contents.match(/<text:changed-region/g) ?? []
    expect(regions.length).toBeGreaterThanOrEqual(4)
  })

  test('track changes: what is typed and deleted becomes a revision in the file and the draft', async () => {
    const file = join(directory, 'controlado.docx')
    const draft = join(directory, 'controlado.sdoc')
    const editor = session.window.locator('.ProseMirror')
    const indicator = session.window.getByTestId('track-changes-status')

    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('Texto antigo.')
    await expect(indicator).toHaveCount(0)
    await menu(session, 'toggle-track-changes')
    await expect(indicator).toBeVisible()

    // Backspace keeps the full stop as deleted; what is typed afterwards is an insertion.
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type(' e novo')
    await expect(editor.locator('del.revision')).toHaveText('.')
    await expect(editor.locator('ins.revision')).toHaveText(' e novo')

    await stubDialogs(session.app, { save: file, open: file, messageBox: 1 })
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(file, 'word/document.xml')
    expect(body).toMatch(/<w:ins [^>]*w:author="[^"]+"/)
    expect(body).toMatch(/<w:del [^>]*w:author="[^"]+"/)
    expect(body).toContain('<w:delText')
    expect(await entryOf(file, 'word/settings.xml')).toContain('<w:trackRevisions')

    if (await hasSoffice()) {
      const copy = join(directory, 'copia.docx')
      await copyFile(file, copy)
      await promisify(execFile)(
        'soffice',
        ['--headless', '--convert-to', 'odt', '--outdir', directory, copy],
        {
          timeout: 120_000,
        },
      )
      const contents = await entryOf(join(directory, 'copia.odt'), 'content.xml')
      expect((contents.match(/<text:changed-region/g) ?? []).length).toBeGreaterThanOrEqual(2)
    }

    // Reopened, the marks and the switch remain; accepting all cleans the text.
    await menu(session, 'open')
    await expect(editor.locator('ins.revision')).toHaveText(' e novo')
    await expect(editor.locator('del.revision')).toHaveText('.')
    await expect(indicator).toBeVisible()
    await menu(session, 'accept-all-changes')
    await expect(editor.locator('ins.revision, del.revision, [data-revision]')).toHaveCount(0)
    await expect(editor).toContainText('Texto antigo e novo')

    // The draft keeps the switch.
    await stubDialogs(session.app, { save: draft, open: draft })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(JSON.parse(await readFile(draft, 'utf8'))).toMatchObject({ trackChanges: true })
    await menu(session, 'toggle-track-changes')
    await expect(indicator).toHaveCount(0)
    await menu(session, 'open')
    await expect(indicator).toBeVisible()
  })

  test('show: simple markup, no markup and original, with the cursor outside the hidden part', async () => {
    const sourceFile = join(directory, 'revisado.docx')
    const destinationFile = join(directory, 'revisado.pdf')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const deleted = editor.locator('del.revision', { hasText: 'Trecho excluído.' })
    const inserted = editor.locator('ins.revision', { hasText: 'com uma inserção' })
    const second = editor.locator('p', { hasText: 'Segundo parágrafo' })
    await expect(deleted).toBeVisible()

    // Simple markup: deletions and the deleted row vanish, the bar appears.
    await menu(session, 'show-simple-markup')
    await expect(editor).toHaveClass(/revisions-simple/)
    await expect(deleted).toBeHidden()
    await expect(editor.locator('tr', { hasText: 'Linha excluída' })).toBeHidden()
    await expect(inserted).toBeVisible()
    await expect(second).toHaveClass(/revision-changed/)

    // The cursor does not enter hidden deleted text: what is typed at the end and after two arrow
    // presses lands in the visible text.
    await inserted.click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type('!')
    await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.type('Z')
    await expect(second).toContainText('revisadaZ.!', { useInnerText: true })
    await expect(deleted).toHaveText('Trecho excluído.')

    // No markup: the insertion stays as plain text, without a bar.
    await menu(session, 'show-no-markup')
    await expect(deleted).toBeHidden()
    expect(await inserted.evaluate((element) => getComputedStyle(element).textDecorationLine)).toBe('none')
    await expect(editor.locator('.revision-changed')).toHaveCount(0)

    if (await hasPdftotext()) {
      await menu(session, 'export-pdf')
      await expect.poll(() => pdfText(destinationFile), { timeout: 30_000 }).toContain('Ata da reunião')
      const text = await pdfText(destinationFile)
      expect(text).toContain('Linha que fica')
      expect(text).not.toContain('Trecho excluído')
      expect(text).not.toContain('Linha excluída')
    }

    // Original: the insertion vanishes, the deletion comes back, and the editor refuses typing.
    await menu(session, 'show-original')
    await expect(inserted).toBeHidden()
    await expect(deleted).toBeVisible()
    await expect(editor.locator('tr', { hasText: 'Linha excluída' })).toBeVisible()
    await expect(editor).toHaveAttribute('contenteditable', 'false')

    await menu(session, 'show-all-markup')
    await expect(inserted).toBeVisible()
    await expect(editor).toHaveAttribute('contenteditable', 'true')
  })

  test('pagination follows the show mode', async () => {
    const editor = session.window.locator('.ProseMirror')
    await menu(session, 'new-document')
    await editor.click()
    for (let line = 1; line <= 60; line++) {
      await session.window.keyboard.insertText(`Linha ${line}`)
      await session.window.keyboard.press('Enter')
    }
    await session.window.keyboard.insertText('Fica.')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    // Everything but the last line, deleted with tracking on.
    await menu(session, 'toggle-track-changes')
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Control+Shift+Home')
    await session.window.keyboard.press('Delete')
    await expect(editor.locator('del.revision').first()).toBeVisible()
    await expect(session.window.locator('.paper')).toHaveCount(2)

    await menu(session, 'show-no-markup')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await menu(session, 'show-all-markup')
    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('consecutive Backspaces with tracking on come out as a single undo', async () => {
    const editor = session.window.locator('.ProseMirror')
    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('abcdef')
    await menu(session, 'toggle-track-changes')
    await session.window.keyboard.press('End')
    for (let time = 0; time < 3; time++) await session.window.keyboard.press('Backspace')
    await expect(editor.locator('del.revision')).toHaveText('def')

    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('del.revision')).toHaveCount(0)
    await expect(editor).toHaveText('abcdef')
  })

  test('a list pasted at the end of a paragraph enters as a revision, and rejecting all removes it', async () => {
    const editor = session.window.locator('.ProseMirror')
    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('Antes da lista.')
    await menu(session, 'toggle-track-changes')
    await session.window.keyboard.press('End')

    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({ text: 'um\ndois', html: '<ul><li>um</li><li>dois</li></ul>' })
    })
    await session.window.keyboard.press('Control+v')
    await expect(editor.locator('li')).toHaveCount(2)
    await expect(editor.locator('ins.revision')).toHaveText(['um', 'dois'])

    await menu(session, 'reject-all-changes')
    await expect(editor.locator('li')).toHaveCount(0)
    await expect(editor).toHaveText('Antes da lista.')
  })
})
