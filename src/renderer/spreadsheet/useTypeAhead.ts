import { useCallback, useEffect, useMemo, useRef } from 'react'

/**
 * A tecla perdida entre uma célula e a seguinte.
 *
 * Depois do Enter o grid **espera 70 ms fixos** antes de mover o foco para
 * baixo — `RESIZE_INTERVAL + 30` na `keyboard.service` dele, uma pausa para não
 * pular a tela caso a grade tenha sido redimensionada. Quem lança uma coluna de
 * números sem parar entre eles acerta essa janela: a tecla chega enquanto o grid
 * ainda aponta para a célula anterior, e ou vira sufixo dela ou não vira nada.
 * Medido: com 50 ms entre o Enter e a tecla seguinte, `1200` abaixo de `980`
 * virava `200`. Perda silenciosa, que só aparece quando a soma não bate no fim
 * do mês.
 *
 * Então a tecla é guardada enquanto a janela está aberta e devolvida quando o
 * foco chega. O `beforekeydown` é o evento que a própria biblioteca oferece para
 * isso — "use this event to check if it wasn't processed by internal logic".
 */
export interface TypeAhead {
  /** Abre a janela de transição, ao confirmar uma edição. */
  readonly begin: () => void
  /** Fecha a janela e devolve o que ficou guardado, quando o foco chega. */
  readonly settle: () => void
}

/**
 * Quanto tempo a janela fica aberta sem que o foco chegue.
 *
 * É a saída para o commit que não move o foco — confirmar clicando noutra
 * célula, por exemplo. Sem ela, o que fosse digitado depois ficaria guardado
 * para sempre.
 */
const WINDOW_MS = 250

export function useTypeAhead(readOnly: boolean): TypeAhead {
  /** Teclas digitadas durante a janela, ainda sem dono. */
  const typed = useRef<string[]>([])
  /**
   * A última tecla já guardada.
   *
   * O grid tem uma sobreposição de seleção por seção do viewport — dados,
   * coluna congelada, linha congelada — e **cada uma** emite o seu
   * `beforekeydown` para a mesma tecla. Sem isto, um `1` seria guardado nove
   * vezes e devolvido nove vezes.
   */
  const lastHeld = useRef<KeyboardEvent | null>(null)
  const open = useRef(false)
  const timer = useRef<number | null>(null)

  const closeWindow = useCallback(() => {
    open.current = false
    if (timer.current !== null) {
      window.clearTimeout(timer.current)
      timer.current = null
    }
  }, [])

  /**
   * Devolve as teclas guardadas pelo caminho normal do grid.
   *
   * Reemitir `keydown` em vez de escrever na célula é o que mantém uma única
   * definição de "digitar por cima de uma célula" — inclusive o `pending edit`
   * dele, que já cuida das teclas que chegam antes de o editor montar. Uma
   * segunda definição nossa divergiria da dele na primeira atualização.
   */
  const replay = useCallback(() => {
    const held = typed.current
    typed.current = []
    if (readOnly) return

    for (const key of held) {
      document.dispatchEvent(
        new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, composed: true }),
      )
    }
  }, [readOnly])

  // Indireção para `begin` não depender da identidade de `replay`: trocar o
  // temporizador a cada renderização o reiniciaria no meio da janela.
  const replayRef = useRef(replay)
  replayRef.current = replay

  const begin = useCallback(() => {
    open.current = true
    if (timer.current !== null) window.clearTimeout(timer.current)
    timer.current = window.setTimeout(() => {
      open.current = false
      timer.current = null
      replayRef.current()
    }, WINDOW_MS)
  }, [])

  const settle = useCallback(() => {
    closeWindow()
    replayRef.current()
  }, [closeWindow])

  useEffect(() => {
    /**
     * A janela abre já no Enter, e não só quando a gravação volta.
     *
     * Entre uma coisa e outra cabe uma tecla — quem digita rápido a perde. O
     * `.edit-input-wrapper` é como a própria biblioteca reconhece o editor de
     * célula (`isEditInput`), então isto só dispara ao confirmar uma edição,
     * nunca ao apertar Enter numa célula parada.
     */
    const commit = (event: KeyboardEvent): void => {
      if (readOnly || !event.isTrusted) return
      if (event.key !== 'Enter' && event.key !== 'Tab') return
      if (!(event.target instanceof HTMLElement)) return
      if (event.target.closest('.edit-input-wrapper') === null) return
      begin()
    }

    // No `document` porque o `beforekeydown` sobe até lá — e é lá que o próprio
    // grid escuta o `keydown`. Ele emite o aviso e checa a resposta na mesma
    // pilha, então prevenir aqui chega a tempo.
    const hold = (event: Event): void => {
      if (!open.current || readOnly) return

      const original = (event as CustomEvent<{ original: KeyboardEvent }>).detail.original
      // A tecla devolvida não pode ser guardada de novo: seria um laço.
      if (!original.isTrusted) return
      if (original === lastHeld.current) {
        // Já guardada por outra sobreposição: só falta impedir esta de tratá-la.
        event.preventDefault()
        return
      }
      if (original.ctrlKey || original.metaKey || original.altKey) return
      // Só caractere digitável: seta, Enter e Escape continuam do grid, senão
      // ninguém mais navegaria durante a janela.
      if (original.key.length !== 1) return

      // Sem isto o grid trataria a tecla apontando para a célula anterior, que
      // é a outra metade do defeito — o `1200` que vira `9801200`.
      event.preventDefault()
      original.preventDefault()
      lastHeld.current = original
      typed.current.push(original.key)
    }

    document.addEventListener('keydown', commit, true)
    document.addEventListener('beforekeydown', hold)
    return () => {
      document.removeEventListener('keydown', commit, true)
      document.removeEventListener('beforekeydown', hold)
    }
  }, [begin, readOnly])

  return useMemo(() => ({ begin, settle }), [begin, settle])
}
