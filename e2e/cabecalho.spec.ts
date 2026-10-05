/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid, entryOf } from './fixtures.js'

/**
 * A grid header: the corporate one is a table, with the logo in a cell merged across rows, the
 * title beside it and the numbering on the right.
 */
test.describe('cabeçalho em grade', () => {
  let session: Session
  let pasta: string
  let destino: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-cabecalho-'))
    destino = join(pasta, 'salva.docx')
    session = await launch()

    const origem = join(pasta, 'grade.docx')
    await writeFile(origem, await docxWithHeaderGrid())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('a tabela do cabeçalho é desenhada como tabela', async () => {
    const grade = session.window.locator('.band--header .band__grid').first()
    await expect(grade).toBeVisible()
    await expect(grade.locator('tr')).toHaveCount(4)
    await expect(grade.locator('td')).toHaveCount(5)
  })

  test('a célula mesclada cresce em vez de deixar a linha vazia', async () => {
    // In OOXML vertical merging is not height: the top cell says `restart` and the lower one shows
    // as an empty cell. Drawn as a real cell, it would open a blank band under the logo.
    const selo = session.window.locator('.band--header .band__grid td', { hasText: 'Selo' }).first()
    await expect(selo).toHaveAttribute('rowspan', '4')
  })

  test('o corpo desce para debaixo do cabeçalho, sem se encontrar com ele', async () => {
    // The top margin is a floor: with a taller header, Word and LibreOffice push the body down.
    const medidas = await session.window.evaluate(() => {
      const banda = document.querySelector('.band--header') as HTMLElement | null
      const primeira = document.querySelector('.page__content > *') as HTMLElement | null
      if (banda === null || primeira === null) return null
      return {
        fimDaFaixa: banda.getBoundingClientRect().bottom,
        inicio: primeira.getBoundingClientRect().top,
      }
    })

    expect(medidas).not.toBeNull()
    expect(medidas!.inicio).toBeGreaterThanOrEqual(medidas!.fimDaFaixa)
  })

  test('o texto do cabeçalho é digitável e volta para o arquivo', async () => {
    // The header title is editable, and only that piece's `w:t` is rewritten: table, borders and
    // merging stay byte for byte.
    const titulo = session.window
      .locator('.band--header .band__text')
      .filter({ hasText: 'Título do documento' })
      .first()

    await titulo.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título corrigido')

    // Text goes out on `blur`: changing the setup redraws the sheets, and redrawing under someone
    // typing would take the cursor away.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.band--header .band__grid')).toContainText('Título corrigido')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const cabecalho = await entryOf(destino, 'word/header1.xml')
    expect(cabecalho).toContain('Título corrigido')
    expect(cabecalho).not.toContain('Título do documento')

    // The rest of the part was not regenerated: the grid is still there, with the merge this writer
    // could not produce from scratch.
    expect(cabecalho).toContain('w:vMerge w:val="restart"')
    expect(cabecalho).toContain('Chamado 10001')
  })

  test('o que não tem texto próprio no arquivo não recebe o cursor', async () => {
    // The margin around the band still belongs to the body: clicking it must not take the cursor
    // out of the text.
    const editaveis = await session.window.locator('.band--header .band__text').count()
    const pecas = await session.window.locator('.band--header .band__grid td').count()

    expect(editaveis).toBeGreaterThan(0)
    expect(editaveis).toBeLessThan(pecas + 1)
    await expect(session.window.locator('.band--header')).not.toHaveAttribute('aria-hidden', 'true')
  })

  test('a grade ocupa a faixa inteira, e não um dos terços', async () => {
    const larguras = await session.window.evaluate(() => {
      const banda = document.querySelector('.band--header') as HTMLElement | null
      const grade = document.querySelector('.band--header .band__grid') as HTMLElement | null
      if (banda === null || grade === null) return null
      return { banda: banda.getBoundingClientRect().width, grade: grade.getBoundingClientRect().width }
    })

    expect(larguras).not.toBeNull()
    expect(larguras!.grade).toBeGreaterThan(larguras!.banda * 0.95)
  })
})
