import { create } from 'zustand'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import {
  DEFAULT_EDITOR_PREFERENCES,
  type EditorPreferences,
  type EditorPreferencesPatch,
} from '@shared/types.js'

/**
 * As preferências de edição, do lado da tela.
 *
 * Não é aqui que elas moram: o dono é o processo main, que as grava e aplica a
 * ortografia na sessão do Chromium (ver `src/main/preferences.ts`). Esta loja é
 * uma **cópia para desenhar** — o que a barra de ferramentas mostra marcado e o
 * que o editor obedece.
 *
 * Toda mudança vai ao main e volta pelo aviso: nunca se escreve direto no estado
 * local. É o que mantém o item do menu nativo e o botão da barra dizendo a mesma
 * coisa, mesmo quando foi o menu quem mudou.
 */
interface PreferencesState {
  readonly preferences: EditorPreferences
  /** Recebido do main. Não deve ser chamado pela interface. */
  readonly accept: (preferences: EditorPreferences) => void
}

export const usePreferences = create<PreferencesState>((set) => ({
  preferences: DEFAULT_EDITOR_PREFERENCES,
  accept: (preferences) => set({ preferences }),
}))

/** O valor atual, para quem não é componente — as extensões do editor. */
export function currentPreferences(): EditorPreferences {
  return usePreferences.getState().preferences
}

/**
 * Pede ao main para ligar ou desligar algo.
 *
 * O estado local **não** é tocado aqui: ele muda quando o aviso volta. Adiantar-se
 * mostraria a marca ligada antes de o corretor estar ligado, e num erro a tela
 * ficaria mentindo.
 */
export async function setPreference(patch: EditorPreferencesPatch): Promise<void> {
  const result = await window.api.preferences.set(patch)
  if (result.ok) usePreferences.getState().accept(result.data)
}

/**
 * Liga a cópia local ao main: lê o estado guardado e assina os avisos.
 *
 * Devolve a função de cancelamento, para o `App` desmontar direito.
 */
export function watchPreferences(): () => void {
  const stop = window.api.preferences.onChange((payload) => {
    // Validado com o mesmo contrato que o main usou para mandar: é a segunda
    // ponta do zod, e o preload não pode fazê-lo por nós (roda sandboxed, sem
    // pacotes de terceiros).
    const parsed = pushContracts[IpcChannel.PreferencesChanged].safeParse(payload)
    if (parsed.success) usePreferences.getState().accept(parsed.data)
  })

  void window.api.preferences.get({}).then((result) => {
    if (result.ok) usePreferences.getState().accept(result.data)
  })

  return stop
}
