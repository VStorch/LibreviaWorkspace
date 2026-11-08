import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { MenuCommand } from '@shared/types.js'
import { onEditorCommand, type EditorCommand } from '../document/editor-commands.js'
import { useWorkspace } from '../state/workspace.js'
import { executeCommand, registerHandlers } from './execute.js'

describe('executeCommand', () => {
  const heard: EditorCommand[] = []
  const cleanups: (() => void)[] = []

  beforeEach(() => {
    heard.length = 0
    useWorkspace.setState({ readOnly: false })
    cleanups.push(onEditorCommand((command) => heard.push(command)))
  })

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup()
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

  it('hands a command without a handler to the editor', async () => {
    await executeCommand(MenuCommand.TableInsert)

    expect(heard).toEqual([MenuCommand.TableInsert])
  })

  it('blocks a writing command while the document is locked and lets a reading one through', async () => {
    useWorkspace.setState({ readOnly: true })

    await executeCommand(MenuCommand.TableInsert)
    await executeCommand(MenuCommand.FindReplace)

    expect(heard).toEqual([MenuCommand.FindReplace])
  })
})
