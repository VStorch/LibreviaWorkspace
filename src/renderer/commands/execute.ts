import { runsWhileLocked, type CommandId } from '@shared/commands/registry.js'
import { asEditorCommand, emitEditorCommand } from '../document/editor-commands.js'
import { useWorkspace } from '../state/workspace.js'

export interface CommandHandler {
  run(arg?: string): void | Promise<void>
  enabled?(): boolean
}

export type CommandHandlers = Partial<Record<CommandId, CommandHandler>>

interface Registered {
  readonly source: string
  readonly handler: CommandHandler
}

const handlers = new Map<CommandId, Registered>()

export function registerHandlers(source: string, map: CommandHandlers): () => void {
  const entries = Object.entries(map) as [CommandId, CommandHandler][]
  for (const [id, handler] of entries) handlers.set(id, { source, handler })
  return () => {
    for (const [id] of entries) if (handlers.get(id)?.source === source) handlers.delete(id)
  }
}

/** Until the editors register their own handlers, what has none goes to the open document. */
function runInDocument(id: CommandId): void {
  const editorCommand = asEditorCommand(id)
  if (editorCommand === null) return
  if (useWorkspace.getState().readOnly && !runsWhileLocked(id)) return
  emitEditorCommand(editorCommand)
}

export async function executeCommand(id: CommandId, arg?: string): Promise<void> {
  const registered = handlers.get(id)
  if (registered === undefined) return runInDocument(id)
  if (registered.handler.enabled?.() === false) return
  await registered.handler.run(arg)
}
