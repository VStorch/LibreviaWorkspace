import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithSections, entryOf } from './fixtures.js'

/**
 * Seções (M9): cada folha com o papel, a faixa e o número da sua seção.
 *
 * O documento tem retrato (romanos), paisagem (reinicia em 1) e retrato de novo
 * começando em página ímpar — o Word insere uma folha em branco antes dela, já
 * com o papel da seção nova. As duas últimas seções não declaram rodapé, e
 * herdam o da primeira.
 */
test.describe('seções', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-secoes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  async function abrir(): Promise<void> {
    const origem = join(pasta, 'secoes.docx')
    await writeFile(origem, await docxWithSections())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(4)
  }

  test('retrato e paisagem na mesma pilha, com a folha em branco da seção ímpar', async () => {
    await abrir()
    const folhas = await session.window
      .locator('.paper')
      .evaluateAll((papeis) => papeis.map((papel) => papel.getBoundingClientRect()))

    // A segunda folha é mais larga que alta; as outras, retrato.
    expect(folhas[0]!.width).toBeLessThan(folhas[0]!.height)
    expect(folhas[1]!.width).toBeGreaterThan(folhas[1]!.height)
    expect(folhas[2]!.width).toBeLessThan(folhas[2]!.height)
    // Centradas na pilha: a de retrato começa mais à direita que a de paisagem.
    expect(folhas[0]!.left).toBeGreaterThan(folhas[1]!.left)

    // O texto da seção de paisagem ocupa a largura dela, e não a do retrato.
    const paisagem = session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' })
    const caixa = await paisagem.boundingBox()
    expect(caixa!.x).toBeLessThan(folhas[0]!.left)
    expect(caixa!.y).toBeGreaterThan(folhas[1]!.top)
    expect(caixa!.y).toBeLessThan(folhas[1]!.bottom)

    // A terceira folha é a em branco; o retrato outra vez abre a quarta.
    const retrato = await session.window
      .locator('.ProseMirror p', { hasText: 'Retrato outra vez' })
      .boundingBox()
    expect(retrato!.y).toBeGreaterThan(folhas[3]!.top)
  })

  test('o rodapé herdado numera cada seção com o formato e o reinício dela', async () => {
    await abrir()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((faixas) => faixas.map((faixa) => (faixa.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Página 1', 'Página 2', 'Página 3'])
  })

  test('o PDF sai com o papel de cada seção e os mesmos números', async () => {
    test.skip(!(await temPoppler()), 'pdfinfo não instalado')
    await abrir()
    const destino = join(pasta, 'secoes.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'export-pdf')

    await expect.poll(() => tamanhos(destino), { timeout: 30_000 }).toHaveLength(4)
    const [primeira, segunda, terceira, quarta] = await tamanhos(destino)
    expect(primeira).toEqual('retrato')
    expect(segunda).toEqual('paisagem')
    expect(terceira).toEqual('retrato')
    expect(quarta).toEqual('retrato')

    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', destino, '-'])
    expect(stdout).toMatch(/Página i[\s\S]*Folha em paisagem[\s\S]*Página 1[\s\S]*Página 2[\s\S]*Página 3/)
  })

  const orientacoes = (session: Session) =>
    session.window.locator('.paper').evaluateAll((papeis) =>
      papeis.map((papel) => {
        const caixa = papel.getBoundingClientRect()
        return caixa.width > caixa.height ? 'paisagem' : 'retrato'
      }),
    )

  test('inserir quebra de seção, virar só a seção de baixo e excluir a quebra', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Em retrato.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Em paisagem.')
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-section-next-page')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    // O cursor está na seção de baixo: "nesta seção" vira só ela.
    await session.window.locator('.ProseMirror p', { hasText: 'Em paisagem.' }).click()
    await menu(session, 'page-setup')
    const painel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    await painel.getByLabel('Orientação').selectOption('landscape')
    await expect(painel.getByLabel('Nesta seção')).toBeChecked()
    await painel.getByRole('button', { name: 'Aplicar' }).click()
    await expect.poll(() => orientacoes(session)).toEqual(['retrato', 'paisagem'])

    const destino = join(pasta, 'duas-secoes.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/<w:sectPr/g)).toHaveLength(2)
    expect(corpo).toMatch(/<w:p>(?:(?!<\/w:p>).)*<w:sectPr(?:(?!<\/w:sectPr>).)*w:h="16838"/)
    expect(corpo).toMatch(/<w:sectPr(?:(?!<\/w:sectPr>).)*w:orient="landscape"(?:(?!<w:sectPr).)*<\/w:body>/)

    // Excluída a quebra, o trecho de cima assume o formato do de baixo, como no Word.
    await menu(session, 'delete-section-break')
    await expect.poll(() => orientacoes(session)).toEqual(['paisagem'])
  })

  test('desvincular o rodapé da seção de paisagem dá a ela um rodapé próprio', async () => {
    await abrir()
    await session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' }).click()
    await menu(session, 'page-setup')
    const painel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    const vinculo = painel.getByLabel('Rodapé: vincular ao anterior')
    await expect(vinculo).toBeChecked()
    await vinculo.uncheck()
    await painel.getByRole('button', { name: 'Aplicar' }).click()

    // A cópia é da seção de paisagem: editá-la não muda o rodapé da primeira.
    const peca = session.window.locator('.band--footer .band__text').nth(1)
    await peca.click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Anexo — ')
    await session.window.locator('.ProseMirror').click()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((faixas) => faixas.map((faixa) => (faixa.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Anexo — Página 1', 'Anexo — Página 2', 'Anexo — Página 3'])

    const destino = join(pasta, 'desvinculado.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(destino, 'word/footer1.xml')).not.toContain('Anexo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/<w:footerReference/g)).toHaveLength(2)
  })
})

async function temPoppler(): Promise<boolean> {
  try {
    await promisify(execFile)('pdfinfo', ['-v'])
    return true
  } catch {
    return false
  }
}

/** A orientação de cada folha do PDF, pelo tamanho que o `pdfinfo` declara. */
async function tamanhos(caminho: string): Promise<string[]> {
  try {
    const { stdout } = await promisify(execFile)('pdfinfo', ['-f', '1', '-l', '99', caminho])
    return [...stdout.matchAll(/Page\s+\d+ size:\s+([\d.]+) x ([\d.]+)/g)].map((medida) =>
      Number(medida[1]) > Number(medida[2]) ? 'paisagem' : 'retrato',
    )
  } catch {
    return []
  }
}
