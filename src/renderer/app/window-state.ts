import { useEffect } from 'react'
import type { MenuSnapshot } from '@shared/ipc.js'
import { DocumentKind, MenuCommand } from '@shared/types.js'
import { buildWindowTitle } from '@services/file/formats.js'
import { onHandlersChanged, unavailableCommands } from '../commands/execute.js'
import { REVISION_VIEW_COMMANDS } from '../state/revision-view.js'
import { dirtyTabs } from '../state/shell.js'
import { useTabs } from '../state/tabs.js'
import type { WorkspaceState } from '../state/types.js'
import { useWorkspace } from '../state/workspace.js'
import { t } from '../i18n.js'

type MenuSource = Pick<WorkspaceState, 'file' | 'readOnly' | 'workbook' | 'trackChanges' | 'revisionView'>

export function menuSnapshot(state: MenuSource, disabled: readonly MenuCommand[] = []): MenuSnapshot {
  const kind = state.file?.kind ?? null
  // It belongs to the document: off with a spreadsheet open.
  const trackChanges = state.workbook === null && state.trackChanges === true
  return {
    activeKind: kind,
    editable: kind !== null && kind !== DocumentKind.Pdf && !state.readOnly,
    checked: [
      ...(trackChanges ? [MenuCommand.ToggleTrackChanges] : []),
      REVISION_VIEW_COMMANDS[state.revisionView],
    ],
    disabled: [...disabled],
  }
}

function openPaths(): string[] {
  return useTabs.getState().tabs.flatMap(({ store }) => {
    const path = store.getState().file?.path ?? null
    return path === null ? [] : [path]
  })
}

/** The window title and the state main's menu shows, both from the active tab. */
export function useWindowStateSync(): void {
  useEffect(() => {
    // Only when something shown changes, not on every key press.
    let last = ''

    const sync = (): void => {
      const state = useWorkspace.getState()
      const untitled = t('shell.file.untitled')
      const title = state.file?.name ?? untitled
      const { isDirty } = state
      const payload = {
        title,
        isDirty,
        dirty: dirtyTabs(),
        commands: menuSnapshot(state, unavailableCommands()),
        paths: openPaths(),
      }
      const key = JSON.stringify(payload)
      if (key === last) return

      last = key
      void window.api.window.setState(payload)
      document.title = buildWindowTitle(state.file?.name ?? null, isDirty, 'Librevia', untitled)
    }

    sync()
    const stopTabs = watchEveryTab(sync)
    const stopHandlers = onHandlersChanged(sync)
    return () => {
      stopTabs()
      stopHandlers()
    }
  }, [])
}

/** Hidden tabs too: one that becomes modified out of sight must still stop a quit. */
export function watchEveryTab(changed: () => void): () => void {
  let watched: (() => void)[] = []
  const watchStores = (): void => {
    for (const stop of watched) stop()
    watched = useTabs.getState().tabs.map((tab) => tab.store.subscribe(changed))
  }

  watchStores()
  const stopTabs = useTabs.subscribe(() => {
    watchStores()
    changed()
  })
  return () => {
    stopTabs()
    for (const stop of watched) stop()
  }
}
