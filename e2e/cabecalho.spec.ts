/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid } from './fixtures.js'

/**
 * Cabeçalho em grade.
 *
 * A faixa de três colunas dá conta de texto, e é o que quase todo cabeçalho
 * corporativo precisa. A outra metade é uma tabela: logotipo numa célula
 * mesclada por várias linhas, título ao lado, numeração à direita. Espalhada
 * pelos terços, ela virava uma fileira de palavras que ainda transbordava sobre
 * a primeira linha do texto.
 */
test.describe('cabeçalho em grade', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-cabecalho-'))
    session = await launch()

    const origem = join(pasta, 'grade.docx')
    await writeFile(origem, await docxWithHeaderGrid())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
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
    // No OOXML a mesclagem vertical não é altura: a célula de cima diz
    // `restart` e a de baixo aparece como célula vazia. Desenhada como célula de
    // verdade, ela abriria uma faixa em branco debaixo do logotipo.
    const selo = session.window.locator('.band--header .band__grid td', { hasText: 'Selo' }).first()
    await expect(selo).toHaveAttribute('rowspan', '4')
  })

  test('o corpo desce para debaixo do cabeçalho, sem se encontrar com ele', async () => {
    // A margem de cima é um piso, não uma posição: quando o cabeçalho é mais
    // alto que ela, o Word e o LibreOffice descem o corpo. Sem isso a primeira
    // linha do texto era escrita por cima da última do cabeçalho.
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
