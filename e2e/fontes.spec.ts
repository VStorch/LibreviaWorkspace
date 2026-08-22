/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node,
// que não conhece `document` nem `FontFace`. A referência vale só para este
// arquivo — pôr DOM no `tsconfig` do Node deixaria o processo main achar que
// tem janela.

import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * As fontes empacotadas chegam à tela.
 *
 * O documento pede Calibri, Cambria, Arial e Times New Roman; nenhuma existe num
 * Linux limpo, e o que o sistema põe no lugar tem métrica própria — a linha
 * quebra noutro ponto e a contagem de páginas muda. As substitutas viajam no
 * instalador e são servidas por um esquema próprio (`src/main/fonts.ts`).
 *
 * O teste carrega a fonte **pela URL**, e não pelo nome: quem programa costuma
 * ter as fontes do sistema instaladas, e aí a regra resolve pelo `local()` sem
 * nunca tocar no arquivo empacotado. O caminho que quebra na máquina do usuário
 * é justamente o que o `local()` esconde aqui.
 */
test.describe('fontes empacotadas', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('o esquema serve a fonte que o instalador leva', async () => {
    const resultado = await session.window.evaluate(async () => {
      const face = new FontFace(
        'ProvaDeCarregamento',
        "url('librevia-font://fonts/Carlito-Regular.ttf') format('truetype')",
      )
      const carregada = await face.load()
      return carregada.status
    })

    expect(resultado).toBe('loaded')
  })

  test('as cinco famílias ficam disponíveis pelo nome do documento', async () => {
    const disponiveis = await session.window.evaluate(() =>
      ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'].filter((familia) =>
        document.fonts.check(`12pt '${familia}'`),
      ),
    )

    expect(disponiveis).toEqual(['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'])
  })

  test('o esquema não serve arquivo de fora da pasta de fontes', async () => {
    // O pedido nasce do CSS, e o CSS pode vir de um documento que qualquer um
    // escreveu. `../` normalizado é a forma clássica de sair de uma pasta que se
    // acreditava fechada.
    const vazou = await session.window.evaluate(async () => {
      try {
        const face = new FontFace(
          'Travessia',
          "url('librevia-font://fonts/../../package.json') format('truetype')",
        )
        await face.load()
        return true
      } catch {
        return false
      }
    })

    expect(vazou).toBe(false)
  })
})
