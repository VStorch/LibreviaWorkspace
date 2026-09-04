import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithAnchoredTextBox } from './fixtures.js'

/**
 * A capa é editável.
 *
 * Num modelo de manual, o título e o subtítulo não estão no fluxo: moram em
 * caixas posicionadas na folha, desenhadas fora do `contenteditable`. Quem
 * abria o documento via o título na tela e não tinha onde clicar — e, antes
 * disso, nem editar o resto podia, porque a forma travava o documento inteiro
 * em somente leitura.
 *
 * O caminho de volta é o que este teste segura: o texto digitado na caixa vira
 * atributo do bloco âncora, atravessa o editor e o gravador, e o `w:txbxContent`
 * do arquivo sai com o texto novo — sem perder a caixa.
 */
test.describe('capa editável', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-capa-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o título da caixa é digitável e volta para o arquivo', async () => {
    const origem = join(folder, 'capa.docx')
    const destino = join(folder, 'salva.docx')
    await writeFile(origem, await docxWithAnchoredTextBox())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    // O documento não abre travado: a forma deixou de ser motivo de cadeado.
    await expect(session.window.locator('.readonly-banner')).toHaveCount(0)

    const caixa = session.window.locator('.paper-float--text').filter({ hasText: 'Título da capa' })
    await expect(caixa).toHaveCount(1)

    await caixa.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título trocado')

    // O texto sai da caixa quando ela perde o foco: só então o atributo do
    // bloco muda, para não redesenhar a folha debaixo do cursor.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.paper-float--text')).toContainText('Título trocado')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const corpo = await corpoDoDocumento(destino)
    expect(corpo).toContain('txbxContent')
    expect(corpo).toContain('Título trocado')
    expect(corpo).not.toContain('Título da capa')
  })
})

/** Conteúdo de `word/document.xml` dentro do `.docx`, sem descompactar em disco. */
async function corpoDoDocumento(caminho: string): Promise<string> {
  const { promisify } = await import('node:util')
  const { inflateRaw } = await import('node:zlib')
  const inflate = promisify(inflateRaw)
  const zip = await readFile(caminho)

  for (let i = 0; i + 30 <= zip.length; i++) {
    if (zip.readUInt32LE(i) !== 0x04034b50) continue

    const metodo = zip.readUInt16LE(i + 8)
    const comprimido = zip.readUInt32LE(i + 18)
    const original = zip.readUInt32LE(i + 22)
    const tamanhoNome = zip.readUInt16LE(i + 26)
    const extra = zip.readUInt16LE(i + 28)
    const nome = zip.subarray(i + 30, i + 30 + tamanhoNome).toString('utf8')
    if (nome !== 'word/document.xml') continue

    const fim = i + 30 + tamanhoNome + extra + (comprimido > 0 ? comprimido : original)
    const dados = zip.subarray(i + 30 + tamanhoNome + extra, fim)
    return (metodo === 0 ? dados : await inflate(dados)).toString('utf8')
  }

  throw new Error(`word/document.xml não encontrado em ${caminho}`)
}
