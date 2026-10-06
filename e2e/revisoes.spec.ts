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
test.describe('revisões', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-revisoes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('aceitar tudo, salvar e reabrir: nenhuma revisão sobra', async () => {
    const origem = join(pasta, 'revisado.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
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

    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).not.toMatch(/<w:(ins|del|moveFrom|moveTo)[ >]/)
    expect(corpo).toContain('com uma inserção revisada.')

    await menu(session, 'open')
    await expect(editor).toContainText('com uma inserção revisada.')
    await expect(editor.locator('ins.revision, del.revision, [data-revision]')).toHaveCount(0)
  })

  test('rejeitar uma exclusão devolve o texto, e desfazer a devolve', async () => {
    const origem = join(pasta, 'revisado.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
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

    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).not.toContain('<w:delText xml:space="preserve"> Trecho excluído.')
    expect(corpo).toContain('Trecho excluído.')
    // The same paragraph's insertion is still a revision, with the id it had.
    expect(corpo).toMatch(/<w:ins [^>]*w:id="1"/)
  })

  test('o LibreOffice lê as revisões que o arquivo gravado leva', async () => {
    test.skip(!(await hasSoffice()), 'sem LibreOffice nesta máquina')
    const origem = join(pasta, 'revisado.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    // Editing the revised paragraph: it is rewritten, and revisions come from the marks.
    const editor = session.window.locator('.ProseMirror')
    await editor.getByText('Segundo parágrafo,').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Novo: ')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const copia = join(pasta, 'copia.docx')
    await copyFile(origem, copia)
    await promisify(execFile)('soffice', ['--headless', '--convert-to', 'odt', '--outdir', pasta, copia], {
      timeout: 120_000,
    })
    const conteudo = await entryOf(join(pasta, 'copia.odt'), 'content.xml')
    // Insertion, deletion, paragraph mark, move (both sides) and row.
    const regioes = conteudo.match(/<text:changed-region/g) ?? []
    expect(regioes.length).toBeGreaterThanOrEqual(4)
  })

  test('controlar alterações: o que se digita e apaga vira revisão no arquivo e no rascunho', async () => {
    const arquivo = join(pasta, 'controlado.docx')
    const rascunho = join(pasta, 'controlado.sdoc')
    const editor = session.window.locator('.ProseMirror')
    const indicador = session.window.getByTestId('track-changes-status')

    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('Texto antigo.')
    await expect(indicador).toHaveCount(0)
    await menu(session, 'toggle-track-changes')
    await expect(indicador).toBeVisible()

    // Backspace keeps the full stop as deleted; what is typed afterwards is an insertion.
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type(' e novo')
    await expect(editor.locator('del.revision')).toHaveText('.')
    await expect(editor.locator('ins.revision')).toHaveText(' e novo')

    await stubDialogs(session.app, { save: arquivo, open: arquivo, messageBox: 1 })
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const corpo = await entryOf(arquivo, 'word/document.xml')
    expect(corpo).toMatch(/<w:ins [^>]*w:author="[^"]+"/)
    expect(corpo).toMatch(/<w:del [^>]*w:author="[^"]+"/)
    expect(corpo).toContain('<w:delText')
    expect(await entryOf(arquivo, 'word/settings.xml')).toContain('<w:trackRevisions')

    if (await hasSoffice()) {
      const copia = join(pasta, 'copia.docx')
      await copyFile(arquivo, copia)
      await promisify(execFile)('soffice', ['--headless', '--convert-to', 'odt', '--outdir', pasta, copia], {
        timeout: 120_000,
      })
      const conteudo = await entryOf(join(pasta, 'copia.odt'), 'content.xml')
      expect((conteudo.match(/<text:changed-region/g) ?? []).length).toBeGreaterThanOrEqual(2)
    }

    // Reopened, the marks and the switch remain; accepting all cleans the text.
    await menu(session, 'open')
    await expect(editor.locator('ins.revision')).toHaveText(' e novo')
    await expect(editor.locator('del.revision')).toHaveText('.')
    await expect(indicador).toBeVisible()
    await menu(session, 'accept-all-changes')
    await expect(editor.locator('ins.revision, del.revision, [data-revision]')).toHaveCount(0)
    await expect(editor).toContainText('Texto antigo e novo')

    // The draft keeps the switch.
    await stubDialogs(session.app, { save: rascunho, open: rascunho })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(JSON.parse(await readFile(rascunho, 'utf8'))).toMatchObject({ trackChanges: true })
    await menu(session, 'toggle-track-changes')
    await expect(indicador).toHaveCount(0)
    await menu(session, 'open')
    await expect(indicador).toBeVisible()
  })

  test('mostrar: marcação simples, sem marcação e original, com o cursor fora do escondido', async () => {
    const origem = join(pasta, 'revisado.docx')
    const destino = join(pasta, 'revisado.pdf')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const excluido = editor.locator('del.revision', { hasText: 'Trecho excluído.' })
    const inserido = editor.locator('ins.revision', { hasText: 'com uma inserção' })
    const segundo = editor.locator('p', { hasText: 'Segundo parágrafo' })
    await expect(excluido).toBeVisible()

    // Simple markup: deletions and the deleted row vanish, the bar appears.
    await menu(session, 'show-simple-markup')
    await expect(editor).toHaveClass(/revisions-simple/)
    await expect(excluido).toBeHidden()
    await expect(editor.locator('tr', { hasText: 'Linha excluída' })).toBeHidden()
    await expect(inserido).toBeVisible()
    await expect(segundo).toHaveClass(/revision-changed/)

    // The cursor does not enter hidden deleted text: what is typed at the end and after two arrow
    // presses lands in the visible text.
    await inserido.click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type('!')
    await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.type('Z')
    await expect(segundo).toContainText('revisadaZ.!', { useInnerText: true })
    await expect(excluido).toHaveText('Trecho excluído.')

    // No markup: the insertion stays as plain text, without a bar.
    await menu(session, 'show-no-markup')
    await expect(excluido).toBeHidden()
    expect(await inserido.evaluate((element) => getComputedStyle(element).textDecorationLine)).toBe('none')
    await expect(editor.locator('.revision-changed')).toHaveCount(0)

    if (await hasPdftotext()) {
      await menu(session, 'export-pdf')
      await expect.poll(() => pdfText(destino), { timeout: 30_000 }).toContain('Ata da reunião')
      const texto = await pdfText(destino)
      expect(texto).toContain('Linha que fica')
      expect(texto).not.toContain('Trecho excluído')
      expect(texto).not.toContain('Linha excluída')
    }

    // Original: the insertion vanishes, the deletion comes back, and the editor refuses typing.
    await menu(session, 'show-original')
    await expect(inserido).toBeHidden()
    await expect(excluido).toBeVisible()
    await expect(editor.locator('tr', { hasText: 'Linha excluída' })).toBeVisible()
    await expect(editor).toHaveAttribute('contenteditable', 'false')

    await menu(session, 'show-all-markup')
    await expect(inserido).toBeVisible()
    await expect(editor).toHaveAttribute('contenteditable', 'true')
  })

  test('a paginação acompanha o modo de mostrar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await menu(session, 'new-document')
    await editor.click()
    for (let linha = 1; linha <= 60; linha++) {
      await session.window.keyboard.insertText(`Linha ${linha}`)
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

  test('Backspaces seguidos com o controle ligado saem num desfazer só', async () => {
    const editor = session.window.locator('.ProseMirror')
    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('abcdef')
    await menu(session, 'toggle-track-changes')
    await session.window.keyboard.press('End')
    for (let vez = 0; vez < 3; vez++) await session.window.keyboard.press('Backspace')
    await expect(editor.locator('del.revision')).toHaveText('def')

    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('del.revision')).toHaveCount(0)
    await expect(editor).toHaveText('abcdef')
  })
})
