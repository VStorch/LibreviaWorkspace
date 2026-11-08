import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind, MenuCommand } from '@shared/types.js'
import { EDITOR_COMMANDS, type EditorCommand } from '../document/editor-commands.js'
import { useTabs } from '../state/tabs.js'
import { createWorkspaceStore, useWorkspace } from '../state/workspace.js'
import {
  editorSource,
  executeCommand,
  onHandlersChanged,
  registerHandlers,
  unavailableCommands,
  type CommandHandlers,
} from './execute.js'

const initialTabs = useTabs.getState()

function editorOf(tabId: string, heard: EditorCommand[]): () => void {
  const map: CommandHandlers = Object.fromEntries(
    EDITOR_COMMANDS.map((command) => [command, { run: () => void heard.push(command) }]),
  )
  return registerHandlers(editorSource(tabId), map)
}

describe('executeCommand', () => {
  const heard: EditorCommand[] = []
  const cleanups: (() => void)[] = []

  beforeEach(() => {
    heard.length = 0
    useWorkspace.setState({ readOnly: false, file: null })
    cleanups.push(editorOf(initialTabs.activeId, heard))
  })

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup()
    useTabs.setState(initialTabs, true)
  })

  it('runs the registered handler with the argument', async () => {
    const run = vi.fn()
    cleanups.push(registerHandlers('test', { [MenuCommand.OpenRecent]: { run } }))

    await executeCommand(MenuCommand.OpenRecent, '/tmp/a.docx')

    expect(run).toHaveBeenCalledWith('/tmp/a.docx')
  })

  it('stops reaching the handler once it is unregistered', async () => {
    const run = vi.fn()
    registerHandlers('test', { [MenuCommand.Save]: { run } })()

    await executeCommand(MenuCommand.Save)

    expect(run).not.toHaveBeenCalled()
  })

  it('skips a handler that says it is disabled', async () => {
    const run = vi.fn()
    cleanups.push(registerHandlers('test', { [MenuCommand.Save]: { run, enabled: () => false } }))

    await executeCommand(MenuCommand.Save)

    expect(run).not.toHaveBeenCalled()
  })

  it('hands an editor command to the active tab editor', async () => {
    await executeCommand(MenuCommand.TableInsert)

    expect(heard).toEqual([MenuCommand.TableInsert])
  })

  it('reaches only the active tab editor, whichever registered last', async () => {
    const second: EditorCommand[] = []
    useTabs.setState({ tabs: [...initialTabs.tabs, { id: 'tab-2', store: createWorkspaceStore() }] })
    cleanups.push(editorOf('tab-2', second))

    await executeCommand(MenuCommand.TableInsert)
    useTabs.setState({ activeId: 'tab-2' })
    await executeCommand(MenuCommand.FindReplace)

    expect(heard).toEqual([MenuCommand.TableInsert])
    expect(second).toEqual([MenuCommand.FindReplace])
  })

  it('falls back to the earlier handler once the newer one is released', async () => {
    const older = vi.fn()
    const newer = vi.fn()
    cleanups.push(registerHandlers('older', { [MenuCommand.Save]: { run: older } }))
    registerHandlers('newer', { [MenuCommand.Save]: { run: newer } })()

    await executeCommand(MenuCommand.Save)

    expect(older).toHaveBeenCalledOnce()
    expect(newer).not.toHaveBeenCalled()
  })

  it('blocks a writing command while the document is locked and lets a reading one through', async () => {
    useWorkspace.setState({ readOnly: true })

    await executeCommand(MenuCommand.TableInsert)
    await executeCommand(MenuCommand.FindReplace)

    expect(heard).toEqual([MenuCommand.FindReplace])
  })
  it('leaves the lock to the text: a locked spreadsheet still opens its properties', async () => {
    const run = vi.fn()
    cleanups.push(
      registerHandlers(editorSource(initialTabs.activeId), { [MenuCommand.DocumentProperties]: { run } }),
    )
    useWorkspace.setState({
      readOnly: true,
      file: { path: '/docs/contas.xlsx', name: 'contas.xlsx', kind: DocumentKind.Spreadsheet },
    })

    await executeCommand(MenuCommand.DocumentProperties)

    expect(run).toHaveBeenCalledOnce()
  })
})

describe('unavailableCommands', () => {
  const cleanups: (() => void)[] = []
  const shell = (): CommandHandlers => ({
    [MenuCommand.Open]: { run: vi.fn() },
    [MenuCommand.Save]: { run: vi.fn() },
    [MenuCommand.ExportWord]: { run: vi.fn() },
  })

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup()
    useTabs.setState(initialTabs, true)
    useWorkspace.setState({ readOnly: false, file: null })
  })

  it('greys out on Home what needs a file and what no view on screen answers', () => {
    cleanups.push(registerHandlers('shell', shell()))
    useWorkspace.setState({ file: null })

    const greyed = unavailableCommands()

    expect(greyed).toEqual(expect.arrayContaining([MenuCommand.Save, MenuCommand.TableInsert]))
    expect(greyed).not.toContain(MenuCommand.Open)
  })

  it('greys out on a spreadsheet the commands only a document has', () => {
    cleanups.push(registerHandlers('shell', shell()))
    useWorkspace.setState({ file: { path: null, name: 'Planilha.ssheet', kind: DocumentKind.Spreadsheet } })

    const greyed = unavailableCommands()

    expect(greyed).toEqual(expect.arrayContaining([MenuCommand.ExportWord, MenuCommand.InsertComment]))
    expect(greyed).not.toContain(MenuCommand.Save)
  })

  it('greys out on a locked document what would edit it, and leaves what only reads', () => {
    cleanups.push(editorOf(initialTabs.activeId, []))
    useWorkspace.setState({
      readOnly: true,
      file: { path: '/docs/a.docx', name: 'a.docx', kind: DocumentKind.Document },
    })

    const greyed = unavailableCommands()

    expect(greyed).toContain(MenuCommand.TableInsert)
    expect(greyed).not.toContain(MenuCommand.FindReplace)
  })

  it('tells once when a burst of registrations changes what answers', async () => {
    const changed = vi.fn()
    cleanups.push(onHandlersChanged(changed))

    const unregister = registerHandlers('a', shell())
    cleanups.push(registerHandlers('b', shell()))
    unregister()
    await Promise.resolve()

    expect(changed).toHaveBeenCalledTimes(1)
  })
})
