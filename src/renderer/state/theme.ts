import { useEffect } from 'react'
import type { ResolvedTheme } from '@shared/types.js'
import { usePreferences } from './preferences.js'

/**
 * O tema que a tela desenha, já resolvido.
 *
 * `system` não chega até aqui resolvido do main, e é de propósito: o main
 * escreve a escolha em `nativeTheme.themeSource`, e o Chromium faz a consulta
 * de mídia abaixo responder de acordo. Então a pergunta "claro ou escuro?" tem
 * uma resposta só, vinda do navegador, em vez de duas — uma calculada no main e
 * outra observada aqui, que é como os dois lados passam a discordar.
 *
 * O ganho concreto: em `system`, trocar o tema do sistema operacional com o
 * aplicativo aberto redesenha a tela sem o main mandar nada.
 */
const DARK_QUERY = '(prefers-color-scheme: dark)'

/**
 * Põe `data-theme` na raiz do documento e o mantém em dia.
 *
 * Um atributo, e não uma classe: o CSS já usa `:root[data-theme='dark']` para a
 * escolha explícita e a consulta de mídia para a automática, e os dois precisam
 * do mesmo seletor de raiz.
 *
 * Quem desenha não é este módulo — é o CSS. Aqui só se diz qual dos dois
 * conjuntos de variáveis vale.
 */
export function useTheme(): ResolvedTheme {
  const theme = usePreferences((state) => state.preferences.theme)

  useEffect(() => {
    const media = window.matchMedia(DARK_QUERY)

    const apply = (): void => {
      const resolved: ResolvedTheme = media.matches ? 'dark' : 'light'
      document.documentElement.setAttribute('data-theme', resolved)
      // `color-scheme` é o que faz o Chromium desenhar as barras de rolagem, os
      // campos e os menus nativos na cor certa. Sem isto o papel fica escuro e
      // a barra de rolagem ao lado dele continua branca.
      document.documentElement.style.colorScheme = resolved
    }

    apply()
    media.addEventListener('change', apply)
    return () => media.removeEventListener('change', apply)
    // `theme` é dependência mesmo sem ser lido no efeito: é a troca dele que
    // muda o `themeSource` no main, e reexecutar aqui é o que garante o
    // atributo escrito no mesmo quadro em vez de no evento seguinte.
  }, [theme])

  return window.matchMedia(DARK_QUERY).matches ? 'dark' : 'light'
}
