import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFile, mkdir, mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { entryOf } from './fixtures.js'

/**
 * Word templates: File → New from template…, opening a `.dotx` and saving as template.
 *
 * A document created from a template is new and untitled: "save" asks for the destination, the
 * `.docx` goes out with the document label, and the template stays as it was.
 */
test.describe('modelos', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-modelos-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const texto = (s: Session) => s.window.locator('.pages__column .ProseMirror')
  const DOCUMENTO = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
  const MODELO = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'

  test('criar a partir do modelo embutido, digitar e salvar pergunta o destino', async () => {
    const modelo = resolve('resources/templates/relatorio.dotx')
    const antes = await hash(modelo)
    const destino = join(pasta, 'relatorio-final.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })

    await menu(session, 'new-from-template')
    const galeria = session.window.getByRole('dialog', { name: 'Novo a partir de modelo' })
    await galeria.getByRole('option', { name: /Relatório com capa e sumário/ }).click()
    await galeria.getByRole('button', { name: 'Criar' }).click()
    await expect(galeria).toHaveCount(0)

    // The template's styles and footer arrived: headings are `h1`, and the band carries the page
    // number.
    await expect(texto(session)).toContainText('[Título do relatório]')
    await expect(texto(session).locator('h1').first()).toHaveText('1 Introdução')
    await expect(session.window.locator('.band--footer').last()).toContainText('Página')

    await texto(session).locator('h1').last().click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' final')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect.poll(() => existe(destino)).toBe(true)

    expect(await hash(modelo)).toBe(antes)
    const tipos = await entryOf(destino, '[Content_Types].xml')
    expect(tipos).toContain(DOCUMENTO)
    expect(tipos).not.toContain(MODELO)
    expect(await entryOf(destino, 'word/document.xml')).toContain('3 Conclusão final')

    // LibreOffice opens the result.
    if (await temSoffice()) {
      const copia = join(pasta, 'lo.docx')
      await copyFile(destino, copia)
      await promisify(execFile)('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', pasta, copia], {
        timeout: 120_000,
      })
      expect(await existe(join(pasta, 'lo.pdf'))).toBe(true)
    }

    // And so does the app itself, now as a regular document.
    await stubDialogs(session.app, { open: destino })
    await menu(session, 'open')
    await expect(texto(session)).toContainText('3 Conclusão final')
    await expect(texto(session).locator('h1').first()).toHaveText('1 Introdução')
  })

  test('salvar como modelo e reabri-lo cria um documento novo', async () => {
    const modelo = join(pasta, 'meu.dotx')
    await stubDialogs(session.app, { save: modelo, messageBox: 1 })

    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Modelo próprio')
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const tipos = await entryOf(modelo, '[Content_Types].xml')
    expect(tipos).toContain(MODELO)
    expect(tipos).not.toContain(DOCUMENTO)
    const antes = await hash(modelo)

    // Reopened, the template becomes a new document: "save" asks for the destination.
    const derivado = join(pasta, 'derivado.docx')
    await stubDialogs(session.app, { open: modelo, save: derivado })
    await menu(session, 'open')
    await expect(texto(session)).toContainText('Modelo próprio')
    await texto(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' — cópia')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await hash(modelo)).toBe(antes)
    expect(await entryOf(derivado, '[Content_Types].xml')).toContain(DOCUMENTO)
    expect(await entryOf(derivado, 'word/document.xml')).toContain('Modelo próprio — cópia')
  })

  test('a galeria mostra os modelos da pasta do usuário', async () => {
    const pastaModelos = join(session.userData, 'Modelos')
    await mkdir(pastaModelos, { recursive: true })
    await copyFile(resolve('resources/templates/carta.dotx'), join(pastaModelos, 'Proposta.dotx'))

    await menu(session, 'new-from-template')
    const galeria = session.window.getByRole('dialog', { name: 'Novo a partir de modelo' })
    await galeria.getByRole('option', { name: /Proposta/ }).dblclick()
    await expect(galeria).toHaveCount(0)
    await expect(texto(session)).toContainText('[Seu nome]')
  })
})

async function hash(path: string): Promise<string> {
  return createHash('sha256')
    .update(await readFile(path))
    .digest('hex')
}

async function existe(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}

async function temSoffice(): Promise<boolean> {
  try {
    await promisify(execFile)('soffice', ['--version'])
    return true
  } catch {
    return false
  }
}
