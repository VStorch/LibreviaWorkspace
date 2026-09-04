/// <reference lib="dom" />
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithSpacingOnBothSides } from './fixtures.js'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * O espaço entre parágrafos soma, como no Word.
 *
 * O CSS funde a margem de baixo de um bloco com a de cima do seguinte e fica
 * com a maior; o Word e o LibreOffice somam as duas. Num documento em que cada
 * parágrafo pede 14 pt depois e o seguinte 14 pt antes, é meia linha por junta
 * — e ela se acumula até a folha cortar noutro lugar.
 */
test.describe('espaçamento entre parágrafos', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-espaco-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o espaço depois e o espaço antes se somam na junta', async () => {
    const origem = join(folder, 'espaco.docx')
    await writeFile(origem, await docxWithSpacingOnBothSides())

    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Antes da junta')

    const distancia = await session.window.evaluate(() => {
      const blocos = document.querySelectorAll('.ProseMirror > p')
      const antes = blocos[0]!.getBoundingClientRect()
      const depois = blocos[1]!.getBoundingClientRect()
      return depois.top - antes.bottom
    })

    // 14,15 pt de cada lado, em pixels de CSS: 2 × 14,15 × 96/72.
    expect(distancia).toBeCloseTo(2 * 14.15 * (96 / 72), 0)
  })
})
