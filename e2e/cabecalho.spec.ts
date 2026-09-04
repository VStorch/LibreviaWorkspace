/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid, entryOf } from './fixtures.js'

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

  test('o texto do cabeçalho é digitável e volta para o arquivo', async () => {
    // A faixa era desenho: o que voltava para o `.docx` era sempre a parte
    // original, e quem via um título errado no cabeçalho não tinha onde clicar.
    // Agora tem — e só o `w:t` daquela peça é reescrito: a tabela, as bordas e
    // a mesclagem seguem byte a byte.
    const titulo = session.window
      .locator('.band--header .band__text')
      .filter({ hasText: 'Título do documento' })
      .first()

    await titulo.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título corrigido')

    // O texto sai no `blur`: mudar a configuração redesenha as folhas, e
    // redesenhar debaixo de quem digita levaria o cursor embora.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.band--header .band__grid')).toContainText('Título corrigido')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const cabecalho = await entryOf(destino, 'word/header1.xml')
    expect(cabecalho).toContain('Título corrigido')
    expect(cabecalho).not.toContain('Título do documento')

    // O resto da parte não foi regenerado: a grade continua lá, com a
    // mesclagem que este escritor não saberia produzir do zero.
    expect(cabecalho).toContain('w:vMerge w:val="restart"')
    expect(cabecalho).toContain('Chamado 10001')
  })

  test('o que não tem texto próprio no arquivo não recebe o cursor', async () => {
    // A margem em volta da faixa continua pertencendo ao corpo: clicar nela não
    // pode tirar o cursor do texto.
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
