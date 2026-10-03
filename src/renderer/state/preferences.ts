import { create } from 'zustand'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import {
  DEFAULT_EDITOR_PREFERENCES,
  type EditorPreferences,
  type EditorPreferencesPatch,
} from '@shared/types.js'

/**
 * Uma **cópia para desenhar**: o dono é o main (`src/main/preferences.ts`). Toda
 * mudança vai ao main e volta pelo aviso, para o menu e a barra dizerem o mesmo.
 */
interface PreferencesState {
  readonly preferences: EditorPreferences
  readonly accept: (preferences: EditorPreferences) => void
}

export const usePreferences = create<PreferencesState>((set) => ({
  preferences: DEFAULT_EDITOR_PREFERENCES,
  accept: (preferences) => set({ preferences }),
}))

/** Para quem não é componente, como as extensões do editor. */
export function currentPreferences(): EditorPreferences {
  return usePreferences.getState().preferences
}

/** O estado local só muda quando o aviso volta: adiantar-se mostraria a marca antes de o corretor ligar. */
export async function setPreference(patch: EditorPreferencesPatch): Promise<void> {
  const result = await window.api.preferences.set(patch)
  if (result.ok) usePreferences.getState().accept(result.data)
}

/** Devolve a função de cancelamento. */
export function watchPreferences(): () => void {
  const stop = window.api.preferences.onChange((payload) => {
    // A segunda ponta do zod, que o preload sandboxed não pode fazer.
    const parsed = pushContracts[IpcChannel.PreferencesChanged].safeParse(payload)
    if (parsed.success) usePreferences.getState().accept(parsed.data)
  })

  void window.api.preferences.get({}).then((result) => {
    if (result.ok) usePreferences.getState().accept(result.data)
  })

  return stop
}
