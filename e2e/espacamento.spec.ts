/// <reference lib="dom" />
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithSpacingOnBothSides } from './fixtures.js'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * O espaço entre parágrafos soma, como no Word e no LibreOffice; o CSS ficaria com a
 * maior margem, meia linha a menos por junta de 14 pt com 14 pt.
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
