import { useEffect } from 'react'
import { setPreference, usePreferences } from './preferences.js'

/**
 * O modo de leitura, do lado da tela.
 *
 * A preferência mora no main, como as outras — é de lá que o item marcado do
 * menu "Exibir" sai. Aqui só se lê a cópia e se pede a troca.
 */
export function useReadingMode(): boolean {
  return usePreferences((state) => state.preferences.readingMode)
}

/** Liga, desliga, alterna. */
export async function setReadingMode(readingMode: boolean): Promise<void> {
  await setPreference({ readingMode })
}

/**
 * Esc sai do modo de leitura.
 *
 * É a única tecla que é preciso saber para não ficar preso, e por isso ela é
 * global: sem barra de ferramentas nem barra de status na tela, um modo que só
 * se desfaz pelo menu é um modo em que a pessoa se sente trancada. O menu
 * continua lá, e `Ctrl+F11` também.
 *
 * Só age quando o modo está ligado. Fora dele, Esc pertence a quem estiver
 * aberto — o painel de localizar, um diálogo — e roubá-lo daqui fecharia o
 * modo de leitura que nem estava aberto e deixaria o diálogo de pé.
 */
export function useLeaveReadingOnEscape(active: boolean): void {
  useEffect(() => {
    if (!active) return undefined

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape' || event.defaultPrevented) return
      event.preventDefault()
      void setReadingMode(false)
    }

    // Na fase de captura: no modo de leitura o editor está travado, mas o
    // ProseMirror ainda recebe as teclas, e um Esc consumido lá dentro nunca
    // chegaria até aqui.
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [active])
}
