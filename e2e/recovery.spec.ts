import { stat } from 'node:fs/promises'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * A promessa da recuperação só vale se for provada contra uma queda de verdade.
 *
 * Um teste que fechasse o aplicativo educadamente provaria outra coisa: o
 * caminho de saída limpa roda handlers, e é justamente o que não acontece quando
 * a máquina desliga ou o processo é morto. Por isso aqui é `SIGKILL`.
 */
test.describe('recuperação depois de uma queda', () => {
  let session: Session

  test.afterEach(async () => {
    await session.close()
  })

  test('devolve o que estava na tela e não escreve em arquivo nenhum', async () => {
    session = await launch()

    await menu(session, 'new-document')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()

    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Ata da reunião de terça')

    // O autosave é por relógio, de oito em oito segundos.
    await expect
      .poll(async () => draftOnDisk(session), { timeout: 20_000, message: 'o rascunho não foi gravado' })
      .toBe(true)

    const { userData } = session
    await session.crash()

    // Mesma pasta de dados: é o que faz a segunda sessão encontrar o rascunho
    // da primeira, como aconteceria com o mesmo usuário na mesma máquina.
    session = await launch({ userData })

    const banner = session.window.locator('.banner--recovery')
    await expect(banner).toBeVisible()
    await expect(banner).toContainText('trabalho não salvo')

    await banner.getByRole('button', { name: 'Recuperar' }).click()

    await expect(session.window.locator('.ProseMirror')).toContainText('Ata da reunião de terça')
    // Recuperado é diferente do que está no disco: marcar como salvo faria o
    // usuário fechar a janela achando que estava tudo guardado.
    await expect(session.window.locator('.statusbar__state')).toContainText('Não salvo')
  })

  test('descartar apaga o rascunho de vez', async () => {
    session = await launch()

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Rascunho a descartar')
    await expect.poll(async () => draftOnDisk(session), { timeout: 20_000 }).toBe(true)

    const { userData } = session
    await session.crash()

    session = await launch({ userData })
    await session.window.locator('.banner--recovery').getByRole('button', { name: 'Descartar' }).click()
    await expect(session.window.locator('.banner--recovery')).toBeHidden()

    expect(await draftOnDisk(session)).toBe(false)
  })

  test('sessão limpa não mostra aviso nenhum', async () => {
    // Um aviso que aparece em toda abertura é um aviso que o usuário aprende a
    // fechar sem ler.
    session = await launch()

    await session.window.waitForTimeout(2000)
    await expect(session.window.locator('.banner--recovery')).toBeHidden()
  })
})

/**
 * O rascunho existe no disco?
 *
 * Conferido de fora, pelo caminho, e não perguntando ao aplicativo: se o teste
 * usasse a mesma API que o código sob teste, um erro no cálculo do caminho
 * passaria despercebido nos dois lados.
 */
async function draftOnDisk(session: Session): Promise<boolean> {
  try {
    await stat(join(session.userData, 'recuperacao', 'rascunho.json'))
    return true
  } catch {
    return false
  }
}
