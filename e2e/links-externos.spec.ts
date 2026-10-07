import { expect, test } from '@playwright/test'
import { launch, type Session } from './app.js'

const ADDRESS = 'https://example.com/relatorio?id=7'

let session: Session

test.beforeEach(async () => {
  session = await launch()
  await session.app.evaluate(({ shell }) => {
    const opened: string[] = []
    Object.assign(globalThis, { openedLinks: opened })
    shell.openExternal = async (url) => {
      opened.push(url)
    }
  })
})

test.afterEach(async () => {
  await session.close()
})

test('um link externo só abre no navegador depois de confirmado, mostrando o endereço', async () => {
  await answerConfirmation(session, 0)

  await session.window.evaluate((url) => window.open(url), ADDRESS)

  await expect.poll(() => openedLinks(session)).toEqual([ADDRESS])
  expect(await askedAbout(session)).toEqual([ADDRESS])
})

test('recusar a confirmação não abre nada', async () => {
  await answerConfirmation(session, 1)

  await session.window.evaluate((url) => window.open(url), ADDRESS)

  await expect.poll(() => askedAbout(session)).toEqual([ADDRESS])
  expect(await openedLinks(session)).toEqual([])
})

test('um endereço fora da lista permitida nem chega a perguntar', async () => {
  await answerConfirmation(session, 0)

  await session.window.evaluate(() => window.open('file:///etc/passwd'))
  await session.window.evaluate((url) => window.open(url), ADDRESS)

  await expect.poll(() => openedLinks(session)).toEqual([ADDRESS])
  expect(await askedAbout(session)).toEqual([ADDRESS])
})

/** Answers every message box with `response`, recording the address each one showed. */
async function answerConfirmation(session: Session, response: number): Promise<void> {
  await session.app.evaluate(({ dialog }, answer) => {
    const asked: string[] = []
    Object.assign(globalThis, { askedLinks: asked })
    dialog.showMessageBox = (async (...args: unknown[]) => {
      asked.push((args.at(-1) as { detail?: string }).detail ?? '')
      return { response: answer, checkboxChecked: false }
    }) as typeof dialog.showMessageBox
  }, response)
}

function askedAbout(session: Session): Promise<string[]> {
  return session.app.evaluate(() => (globalThis as { askedLinks?: string[] }).askedLinks ?? [])
}

function openedLinks(session: Session): Promise<string[]> {
  return session.app.evaluate(() => (globalThis as { openedLinks?: string[] }).openedLinks ?? [])
}
