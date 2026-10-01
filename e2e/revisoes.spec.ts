import { execFile } from 'node:child_process'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { inflateRawSync } from 'node:zlib'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTrackedChange, entryOf } from './fixtures.js'

async function temSoffice(): Promise<boolean> {
  try {
    await promisify(execFile)('soffice', ['--version'])
    return true
  } catch {
    return false
  }
}

/** Uma entrada de um ZIP comum (comprimido), como o LibreOffice grava o `.odt`. */
async function entradaZip(caminho: string, nome: string): Promise<string> {
  const zip = await readFile(caminho)
  // O diretório central no fim: cada entrada diz onde começa o cabeçalho local.
  const fim = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  let posicao = zip.readUInt32LE(fim + 16)
  const total = zip.readUInt16LE(fim + 10)
  for (let index = 0; index < total; index++) {
    const metodo = zip.readUInt16LE(posicao + 10)
    const tamanho = zip.readUInt32LE(posicao + 20)
    const nomeTamanho = zip.readUInt16LE(posicao + 28)
    const extra = zip.readUInt16LE(posicao + 30)
    const comentario = zip.readUInt16LE(posicao + 32)
    const local = zip.readUInt32LE(posicao + 42)
    const entrada = zip.toString('utf8', posicao + 46, posicao + 46 + nomeTamanho)
    if (entrada === nome) {
      const inicio = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28)
      const dados = zip.subarray(inicio, inicio + tamanho)
      return (metodo === 0 ? dados : inflateRawSync(dados)).toString('utf8')
    }
    posicao += 46 + nomeTamanho + extra + comentario
  }
  return ''
}

/**
 * Controle de alterações (M10, fase 1): a revisão aparece na tela, aceitar e
 * rejeitar mudam o texto, e o arquivo gravado diz o mesmo — no Word e no
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
    // A inserção do mesmo parágrafo continua revisão, com o id que tinha.
    expect(corpo).toMatch(/<w:ins [^>]*w:id="1"/)
  })

  test('o LibreOffice lê as revisões que o arquivo gravado leva', async () => {
    test.skip(!(await temSoffice()), 'sem LibreOffice nesta máquina')
    const origem = join(pasta, 'revisado.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    // Editar o parágrafo revisado: ele é reescrito, e as revisões vêm das marcas.
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
    const conteudo = await entradaZip(join(pasta, 'copia.odt'), 'content.xml')
    // Inserção, exclusão, marca de parágrafo, movimentação (dois lados) e linha.
    const regioes = conteudo.match(/<text:changed-region/g) ?? []
    expect(regioes.length).toBeGreaterThanOrEqual(4)
  })
})
