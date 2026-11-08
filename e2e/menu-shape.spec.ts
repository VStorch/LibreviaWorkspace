import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type ElectronApplication } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithoutExtras } from './fixtures.js'

interface MenuShape {
  readonly label: string
  readonly type: string
  readonly role?: string
  readonly accelerator?: string
  readonly registerAccelerator?: boolean
  readonly checked?: boolean
  readonly enabled: boolean
  readonly toolTip?: string
  readonly submenu?: readonly MenuShape[]
}

async function menuShape(app: ElectronApplication): Promise<MenuShape[]> {
  return app.evaluate(({ Menu }) => {
    type Item = Electron.MenuItem
    const shape = (item: Item): MenuShape => ({
      label: item.label,
      type: item.type,
      ...(item.role ? { role: item.role } : {}),
      ...(item.accelerator ? { accelerator: item.accelerator } : {}),
      ...(item.registerAccelerator ? {} : { registerAccelerator: false }),
      ...(item.type === 'checkbox' || item.type === 'radio' ? { checked: item.checked } : {}),
      enabled: item.enabled,
      ...(item.toolTip ? { toolTip: item.toolTip } : {}),
      ...(item.submenu ? { submenu: item.submenu.items.map(shape) } : {}),
    })
    return Menu.getApplicationMenu()?.items.map(shape) ?? []
  })
}

async function setPortuguese(session: Session): Promise<void> {
  await session.window.evaluate(async () => {
    const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } }).api
    await api.preferences.set({ language: 'pt' })
  })
  await expect.poll(async () => (await menuShape(session.app))[0]?.label).toBe('Arquivo')
}

/** With the final newline the formatter asks of a JSON file. */
function asJson(shape: readonly MenuShape[]): string {
  return `${JSON.stringify(shape, null, 2)}\n`
}

function recentLabels(shape: readonly MenuShape[]): string[] {
  const recent = shape[0]?.submenu?.find((item) => item.label === 'Abrir recente')
  return recent?.submenu?.map((item) => item.label) ?? []
}

/** Pins the whole menu, so a refactor of how it is built shows any item it changed. */
test.describe('menu shape', () => {
  // The Redo key and the macOS application menu differ by platform; one recorded shape is enough.
  test.skip(process.platform !== 'linux', 'the snapshot is recorded on Linux')

  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-menu-'))
    session = await launch()
    await setPortuguese(session)
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('on the home screen, with nothing open', async () => {
    expect(recentLabels(await menuShape(session.app))).toEqual(['Nenhum arquivo recente'])
    expect(asJson(await menuShape(session.app))).toMatchSnapshot('home.json')
  })

  test('with a .docx open', async () => {
    const target = join(folder, 'ata.docx')
    await writeFile(target, await docxWithoutExtras())
    await stubDialogs(session.app, { open: target })

    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Ata simples.')
    await expect.poll(async () => recentLabels(await menuShape(session.app))).toContain('ata.docx')

    const shape = asJson(await menuShape(session.app)).replaceAll(folder, '<folder>')
    expect(shape).toMatchSnapshot('docx.json')
  })
})
