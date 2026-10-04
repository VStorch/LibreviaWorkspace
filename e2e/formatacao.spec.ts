import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithVerticalAlignment } from './fixtures.js'

/**
 * Formatação de caractere e de parágrafo pela costura inteira: o botão, o comando
 * do editor e o que o sidecar grava. O sobrescrito muda o sentido ("cm3" não é
 * "cm³"): se ele volta do arquivo, o caminho está de pé.
 */
test.describe('formatação do documento', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-formatacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o sobrescrito do documento aparece na tela e volta para o arquivo', async () => {
    const origem = join(folder, 'formula.docx')
    const destino = join(folder, 'salva.docx')
    await writeFile(origem, await docxWithVerticalAlignment())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('O ocupa')

    // Na tela, `<sup>` e `<sub>`, que o PDF também leva.
    await expect(editor.locator('sup')).toHaveText('3')
    await expect(editor.locator('sub')).toHaveText('2')

    // Editar obriga o gravador a reescrever o parágrafo, onde a perda aconteceria.
    await editor.click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Mexido.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const corpo = await corpoDoDocumento(destino)
    expect(corpo).toContain('w:vertAlign w:val="superscript"')
    expect(corpo).toContain('w:vertAlign w:val="subscript"')

    // Nada no aviso: esta formatação o gravador sabe gravar.
    await expect(session.window.locator('.banner--notice')).toHaveCount(0)
  })

  test('os botões de sobrescrito, caixa alta e versalete marcam o texto', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('marcado')
    await session.window.keyboard.press('Control+a')

    for (const [rotulo, seletor] of [
      ['Sobrescrito', 'sup'],
      ['Caixa alta', 'span[data-caps]'],
      ['Versalete', 'span[data-small-caps]'],
    ] as const) {
      await session.window.getByRole('button', { name: rotulo, exact: true }).click()
      await expect(editor.locator(seletor)).toHaveCount(1)
    }
  })

  test('o diálogo de parágrafo aplica espaçamento e recuo', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo medido.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await expect(dialogo).toBeVisible()

    // Pelo papel: "À esquerda" também é opção do alinhamento, no `label` do seletor.
    await dialogo.getByRole('spinbutton', { name: 'Antes' }).fill('18')
    await dialogo.getByRole('spinbutton', { name: 'Esquerda' }).fill('20')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toBeHidden()

    // Medida de verdade, como a paginação a verá.
    const medidas = await editor
      .locator('p')
      .first()
      .evaluate((elemento) => {
        const estilo = getComputedStyle(elemento)
        return { antes: estilo.marginTop, recuo: estilo.paddingLeft }
      })

    // 18 pt em pixels de CSS, e 20 mm idem.
    expect(Number.parseFloat(medidas.antes)).toBeCloseTo(18 * (96 / 72), 0)
    expect(Number.parseFloat(medidas.recuo)).toBeCloseTo(20 * (96 / 25.4), 0)
  })

  test('a entrelinha de 1,5 linha é 1,5 linha, e não 1,5 de CSS', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo de uma linha e meia.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialogo.getByRole('combobox', { name: 'Entrelinha' }).selectOption('1.5')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()

    // O múltiplo do Word é sobre a altura natural da fonte do estilo: 1,5 linha em
    // Calibri é `line-height: 1.8311`.
    const proporcao = await editor
      .locator('p')
      .first()
      .evaluate((elemento) => {
        const estilo = getComputedStyle(elemento)
        return Number.parseFloat(estilo.lineHeight) / Number.parseFloat(estilo.fontSize)
      })

    expect(proporcao).toBeCloseTo(1.8311, 2)

    // E a barra mostra o número do Word.
    await expect(session.window.getByRole('combobox', { name: 'Espaçamento entre linhas' })).toHaveValue(
      '1.5',
    )
  })

  test('"Aplicar" sem mexer em nada devolve o foco ao texto', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Sem mudança nenhuma.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toBeHidden()

    // "Aplicar" sem nada a mudar devolve o foco ao texto, no quadro seguinte ao fechamento.
    await expect(editor).toBeFocused()
    await session.window.keyboard.type(' Continua.')
    await expect(editor).toContainText('Sem mudança nenhuma. Continua.')
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

    const dados = zip.subarray(i + 30 + tamanhoNome + extra, i + 30 + tamanhoNome + extra + comprimido)
    if (metodo === 0) return dados.subarray(0, original).toString('utf8')
    return (await inflate(dados)).toString('utf8')
  }

  throw new Error('word/document.xml não encontrado no pacote')
}
