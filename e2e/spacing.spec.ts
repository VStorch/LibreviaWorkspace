/// <reference lib="dom" />
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithSpacingOnBothSides } from './fixtures.js'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * Space between paragraphs adds up, as in Word and LibreOffice; CSS would keep the larger margin,
 * half a line less per 14 pt/14 pt joint.
 */
test.describe('paragraph spacing', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-spacing-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('space after and space before add up at the joint', async () => {
    const sourceFile = join(folder, 'espaco.docx')
    await writeFile(sourceFile, await docxWithSpacingOnBothSides())

    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Antes da junta')

    const distance = await session.window.evaluate(() => {
      const blocks = document.querySelectorAll('.ProseMirror > p')
      const before = blocks[0]!.getBoundingClientRect()
      const after = blocks[1]!.getBoundingClientRect()
      return after.top - before.bottom
    })

    // 14.15 pt on each side, in CSS pixels: 2 × 14.15 × 96/72.
    expect(distance).toBeCloseTo(2 * 14.15 * (96 / 72), 0)
  })
})
