import { appliesTo, COMMANDS, runsWhileLocked, type CommandId } from '@shared/commands/registry.js'
import { DocumentKind } from '@shared/types.js'
import { activeWorkspaceStore, useTabs } from '../state/tabs.js'

export interface CommandHandler {
  run(arg?: string): void | Promise<void>
  enabled?(): boolean
}

export type CommandHandlers = Partial<Record<CommandId, CommandHandler>>

interface Registered {
  readonly source: string
  readonly handler: CommandHandler
}

const EDITOR_SOURCE = 'editor:'

/** Reached only while its tab is active, and held back by the read-only lock. */
export function editorSource(tabId: string): string {
  return `${EDITOR_SOURCE}${tabId}`
}

const isEditorSource = (source: string): boolean => source.startsWith(EDITOR_SOURCE)

/** Newest last: the latest registration answers, and releasing it uncovers the one before. */
const handlers = new Map<CommandId, Registered[]>()

const listeners = new Set<() => void>()
let notifying = false

/** Once per burst: switching tabs unregisters one view and registers another in the same pass. */
function handlersChanged(): void {
  if (notifying) return
  notifying = true
  queueMicrotask(() => {
    notifying = false
    for (const listener of listeners) listener()
  })
}

export function onHandlersChanged(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function registerHandlers(source: string, map: CommandHandlers): () => void {
  const added = (Object.entries(map) as [CommandId, CommandHandler][]).map(([id, handler]) => {
    const registered: Registered = { source, handler }
    handlers.set(id, [...(handlers.get(id) ?? []), registered])
    return [id, registered] as const
  })
  handlersChanged()
  return () => {
    for (const [id, registered] of added) {
      const left = (handlers.get(id) ?? []).filter((candidate) => candidate !== registered)
      if (left.length === 0) handlers.delete(id)
      else handlers.set(id, left)
    }
    handlersChanged()
  }
}

function handlerFor(id: CommandId): Registered | undefined {
  const activeEditor = editorSource(useTabs.getState().activeId)
  return (handlers.get(id) ?? []).findLast(({ source }) => !isEditorSource(source) || source === activeEditor)
}

/** The lock belongs to the text: a spreadsheet's and a PDF's views keep their own rules. */
function lockedOut(id: CommandId, { source }: Registered): boolean {
  const { readOnly, file } = activeWorkspaceStore().getState()
  const kind = file?.kind
  const locksText = readOnly && kind !== DocumentKind.Spreadsheet && kind !== DocumentKind.Pdf
  return isEditorSource(source) && locksText && !runsWhileLocked(id)
}

/** What the menu greys out: no view on screen answers it, or the lock holds it back. */
export function unavailableCommands(): CommandId[] {
  const kind = activeWorkspaceStore().getState().file?.kind ?? null
  return (Object.keys(COMMANDS) as CommandId[]).filter((id) => {
    const registered = handlerFor(id)
    return registered === undefined || lockedOut(id, registered) || !appliesTo(id, kind)
  })
}

export async function executeCommand(id: CommandId, arg?: string): Promise<void> {
  const registered = handlerFor(id)
  if (registered === undefined || lockedOut(id, registered)) return
  if (registered.handler.enabled?.() === false) return
  await registered.handler.run(arg)
}
