import { create } from 'zustand'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import {
  DEFAULT_EDITOR_PREFERENCES,
  type EditorPreferences,
  type EditorPreferencesPatch,
} from '@shared/types.js'

/**
 * A **copy for drawing**: main owns them (`src/main/preferences.ts`). Every change goes to main and
 * comes back through the echo, so the menu and the toolbar agree.
 */
interface PreferencesState {
  readonly preferences: EditorPreferences
  readonly accept: (preferences: EditorPreferences) => void
}

export const usePreferences = create<PreferencesState>((set) => ({
  preferences: DEFAULT_EDITOR_PREFERENCES,
  accept: (preferences) => set({ preferences }),
}))

/** For non-components, such as editor extensions. */
export function currentPreferences(): EditorPreferences {
  return usePreferences.getState().preferences
}

/**
 * Local state only changes when the echo comes back: getting ahead would show the check mark before
 * the spellchecker is on.
 */
export async function setPreference(patch: EditorPreferencesPatch): Promise<void> {
  const result = await window.api.preferences.set(patch)
  if (result.ok) usePreferences.getState().accept(result.data)
}

/** Returns the unsubscribe function. */
export function watchPreferences(): () => void {
  const stop = window.api.preferences.onChange((payload) => {
    // The second zod check, which the sandboxed preload cannot do.
    const parsed = pushContracts[IpcChannel.PreferencesChanged].safeParse(payload)
    if (parsed.success) usePreferences.getState().accept(parsed.data)
  })

  void window.api.preferences.get({}).then((result) => {
    if (result.ok) usePreferences.getState().accept(result.data)
  })

  return stop
}
