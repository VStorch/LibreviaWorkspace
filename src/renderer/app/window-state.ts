import { useEffect } from 'react'
import type { MenuSnapshot } from '@shared/ipc.js'
import { DocumentKind, MenuCommand } from '@shared/types.js'
import { buildWindowTitle } from '@services/file/formats.js'
import { REVISION_VIEW_COMMANDS } from '../state/revision-view.js'
import { dirtyTabs } from '../state/shell.js'
import { useTabs } from '../state/tabs.js'
import type { WorkspaceState } from '../state/types.js'
import { useWorkspace } from '../state/workspace.js'
import { t } from '../i18n.js'

type MenuSource = Pick<WorkspaceState, 'file' | 'readOnly' | 'workbook' | 'trackChanges' | 'revisionView'>

export function menuSnapshot(state: MenuSource): MenuSnapshot {
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
    disabled: [],
  }
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
      const payload = { title, isDirty, dirty: dirtyTabs(), commands: menuSnapshot(state) }
      const key = JSON.stringify(payload)
      if (key === last) return

      last = key
      void window.api.window.setState(payload)
      document.title = buildWindowTitle(state.file?.name ?? null, isDirty, 'Librevia', untitled)
    }

    sync()
    const unsubscribe = useWorkspace.subscribe(sync)
    const unsubscribeTabs = useTabs.subscribe(sync)
    return () => {
      unsubscribeTabs()
      unsubscribe()
    }
  }, [])
}
