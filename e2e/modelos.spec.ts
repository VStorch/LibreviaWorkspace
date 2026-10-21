import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFile, mkdir, mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { entryOf } from './fixtures.js'
import { hasSoffice } from './external-tools.js'

/**
 * Word templates: File → New from template…, opening a `.dotx` and saving as template.
 *
 * A document created from a template is new and untitled: "save" asks for the destination, the
 * `.docx` goes out with the document label, and the template stays as it was.
 */
test.describe('templates', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-templates-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const text = (s: Session) => s.window.locator('.pages__column .ProseMirror')
  const DOCUMENT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
  const TEMPLATE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'

  test('creating from the built-in template, typing and saving asks for the destination', async () => {
    const template = resolve('resources/templates/relatorio.dotx')
    const before = await hash(template)
    const destinationFile = join(directory, 'relatorio-final.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })

    await menu(session, 'new-from-template')
    const gallery = session.window.getByRole('dialog', { name: 'Novo a partir de modelo' })
    await gallery.getByRole('option', { name: /Relatório com capa e sumário/ }).click()
    await gallery.getByRole('button', { name: 'Criar' }).click()
    await expect(gallery).toHaveCount(0)

    // The template's styles and footer arrived: headings are `h1`, and the band carries the page
    // number.
    await expect(text(session)).toContainText('[Título do relatório]')
    await expect(text(session).locator('h1').first()).toHaveText('1 Introdução')
    await expect(session.window.locator('.band--footer').last()).toContainText('Página')

    await text(session).locator('h1').last().click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' final')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect.poll(() => exists(destinationFile)).toBe(true)

    expect(await hash(template)).toBe(before)
    const types = await entryOf(destinationFile, '[Content_Types].xml')
    expect(types).toContain(DOCUMENT)
    expect(types).not.toContain(TEMPLATE)
    expect(await entryOf(destinationFile, 'word/document.xml')).toContain('3 Conclusão final')

    // LibreOffice opens the result.
    if (await hasSoffice()) {
      const copy = join(directory, 'lo.docx')
      await copyFile(destinationFile, copy)
      await promisify(execFile)(
        'soffice',
        ['--headless', '--convert-to', 'pdf', '--outdir', directory, copy],
        {
          timeout: 120_000,
        },
      )
      expect(await exists(join(directory, 'lo.pdf'))).toBe(true)
    }

    // And so does the app itself, now as a regular document.
    await stubDialogs(session.app, { open: destinationFile })
    await menu(session, 'open')
    await expect(text(session)).toContainText('3 Conclusão final')
    await expect(text(session).locator('h1').first()).toHaveText('1 Introdução')
  })

  test('saving as template and reopening it creates a new document', async () => {
    const template = join(directory, 'meu.dotx')
    await stubDialogs(session.app, { save: template, messageBox: 1 })

    await menu(session, 'new-document')
    await text(session).click()
    await session.window.keyboard.type('Modelo próprio')
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const types = await entryOf(template, '[Content_Types].xml')
    expect(types).toContain(TEMPLATE)
    expect(types).not.toContain(DOCUMENT)
    const before = await hash(template)

    // Reopened, the template becomes a new document: "save" asks for the destination.
    const derived = join(directory, 'derivado.docx')
    await stubDialogs(session.app, { open: template, save: derived })
    await menu(session, 'open')
    await expect(text(session)).toContainText('Modelo próprio')
    await text(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' — cópia')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await hash(template)).toBe(before)
    expect(await entryOf(derived, '[Content_Types].xml')).toContain(DOCUMENT)
    expect(await entryOf(derived, 'word/document.xml')).toContain('Modelo próprio — cópia')
  })

  test('the gallery shows the templates in the user folder', async () => {
    const templatesDirectory = join(session.userData, 'Modelos')
    await mkdir(templatesDirectory, { recursive: true })
    await copyFile(resolve('resources/templates/carta.dotx'), join(templatesDirectory, 'Proposta.dotx'))

    await menu(session, 'new-from-template')
    const gallery = session.window.getByRole('dialog', { name: 'Novo a partir de modelo' })
    await gallery.getByRole('option', { name: /Proposta/ }).dblclick()
    await expect(gallery).toHaveCount(0)
    await expect(text(session)).toContainText('[Seu nome]')
  })
})

async function hash(path: string): Promise<string> {
  return createHash('sha256')
    .update(await readFile(path))
    .digest('hex')
}

async function exists(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}
