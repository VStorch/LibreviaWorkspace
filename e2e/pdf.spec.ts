import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inflateSync } from 'node:zlib'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/**
 * O PDF sai, e com tinta dentro.
 *
 * A exportação de planilha ficou quatro fases sem funcionar justamente porque
 * falhava **em silêncio**: `buildPrintRequest` devolvia `null`, ninguém era
 * avisado e nenhum arquivo aparecia. Um teste que só verificasse "não deu erro"
 * teria passado o tempo todo — por isso este vai ao disco.
 */
test.describe('exportar PDF', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-pdf-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('documento vira PDF', async () => {
    const target = join(folder, 'ata.pdf')
    await stubDialogs(session.app, { save: target, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Ata da reunião de terça')

    await menu(session, 'export-pdf')
    await expect.poll(() => glyphRuns(target), { timeout: 30_000 }).toBeGreaterThan(0)
  })

  test('planilha vira PDF', async () => {
    const target = join(folder, 'contas.pdf')
    await stubDialogs(session.app, { save: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await expect(
      session.window.locator('revogr-overlay-selection revogr-data [data-rgrow="0"][data-rgcol="0"]').first(),
    ).toBeVisible()

    const input = session.window.locator('.formula-bar__input')
    await input.fill('Aluguel')
    await input.press('Enter')

    await menu(session, 'export-pdf')
    await expect.poll(() => glyphRuns(target), { timeout: 30_000 }).toBeGreaterThan(0)
  })

  test('a linha em branco continua ocupando uma linha no papel', async () => {
    // No editor o parágrafo vazio ocupa uma linha porque o ProseMirror põe um
    // <br> invisível dentro dele. O serializador do papel não põe, e um bloco
    // sem conteúdo tem altura zero — então as folhas eram recortadas pela
    // medida da tela e impressas sem essas linhas. O texto subia, e a primeira
    // linha ia parar debaixo do cabeçalho.
    const target = join(folder, 'linhas.pdf')
    await stubDialogs(session.app, { save: target, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Alfa')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Beta')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Gama')

    await menu(session, 'export-pdf')
    await expect.poll(() => textTops(target), { timeout: 30_000 }).toHaveLength(3)

    // Sem número mágico: o próprio documento dá a medida de um parágrafo, e a
    // linha em branco tem de valer outro tanto.
    const [alfa, beta, gama] = await textTops(target)
    const umParagrafo = beta! - alfa!
    expect(gama! - beta!).toBeGreaterThan(umParagrafo * 1.8)
  })

  test('sem nada aberto, avisa em vez de não fazer nada', async () => {
    await stubDialogs(session.app, { save: join(folder, 'nada.pdf') })

    await menu(session, 'export-pdf')

    await expect(session.window.locator('.banner')).toContainText('Não há nada aberto para imprimir')
  })
})

/**
 * Quantas vezes o PDF manda desenhar texto.
 *
 * Não dá para procurar a palavra: o Chromium embute a fonte como subconjunto e
 * escreve o texto como identificadores de glifo (`<0003> Tj`), que só o mapa da
 * própria fonte traduz. Contar os operadores de texto responde a pergunta que
 * importa aqui — "chegou tinta no papel?" — sem trazer um interpretador de PDF
 * para dentro do teste.
 *
 * Arquivo ausente conta zero, que é exatamente o defeito antigo.
 */
async function glyphRuns(path: string): Promise<number> {
  const bytes = await readFile(path).catch(() => null)
  if (bytes === null) return 0

  let text = ''
  let at = 0
  while (true) {
    const start = bytes.indexOf('stream', at)
    if (start === -1) break

    let from = start + 'stream'.length
    if (bytes[from] === 0x0d) from++
    if (bytes[from] === 0x0a) from++

    const end = bytes.indexOf('endstream', from)
    if (end === -1) break

    try {
      text += inflateSync(bytes.subarray(from, end)).toString('latin1')
    } catch {
      // Fluxo que não é conteúdo comprimido — fonte embutida, imagem. Segue.
    }
    at = end + 'endstream'.length
  }

  return (text.match(/\b(Tj|TJ)\b/g) ?? []).length
}

/**
 * A altura de cada linha de texto desenhada, na ordem em que o PDF as desenha.
 *
 * O texto em si é ilegível — o Chromium embute a fonte como subconjunto e
 * escreve identificadores de glifo. A **posição** não é: cada linha vem
 * precedida de uma matriz de texto `1 0 0 -1 x y Tm`, e é o `y` dela que diz
 * onde a linha caiu. É o bastante para perguntar "quanto desceu daqui para
 * ali" sem trazer um interpretador de PDF para dentro do teste.
 */
async function textTops(path: string): Promise<number[]> {
  const bytes = await readFile(path).catch(() => null)
  if (bytes === null) return []

  let text = ''
  let at = 0
  while (true) {
    const start = bytes.indexOf('stream', at)
    if (start === -1) break

    let from = start + 'stream'.length
    if (bytes[from] === 0x0d) from++
    if (bytes[from] === 0x0a) from++

    const end = bytes.indexOf('endstream', from)
    if (end === -1) break

    try {
      text += inflateSync(bytes.subarray(from, end)).toString('latin1')
    } catch {
      // Fluxo que não é conteúdo comprimido — fonte embutida, imagem. Segue.
    }
    at = end + 'endstream'.length
  }

  const tops: number[] = []
  const matrix = /1 0 0 -1 [\d.]+ ([\d.]+) Tm\s*\n<[^>]*> Tj/g
  let found: RegExpExecArray | null
  while ((found = matrix.exec(text)) !== null) tops.push(Number(found[1]))
  return tops
}
