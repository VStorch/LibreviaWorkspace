import { useEffect } from 'react'
import { Theme, type ResolvedTheme } from '@shared/types.js'
import { usePreferences } from './preferences.js'

const DARK_QUERY = '(prefers-color-scheme: dark)'

/**
 * Claro ou escuro, a partir da escolha e do que o sistema diz.
 *
 * A escolha explícita decide sozinha. **Só** `system` consulta a mídia.
 *
 * A primeira versão disto deixava a consulta decidir sempre, contando com o
 * `nativeTheme.themeSource` que o main escreve para fazer
 * `prefers-color-scheme` mudar junto. É o arranjo mais bonito dos dois e não
 * funcionou: o teste de ponta a ponta pediu o tema escuro e o `data-theme`
 * continuou `light`. Bonito e não verificado perde para direto e conferido —
 * uma escolha explícita agora não depende de o Chromium propagar nada.
 *
 * O `themeSource` continua sendo escrito no main, e continua valendo: é ele que
 * põe menu de contexto, barra de rolagem e janela na cor certa, e é ele que faz
 * `system` responder quando a pessoa troca o tema do sistema operacional sem
 * fechar o aplicativo.
 */
function resolve(theme: Theme, systemPrefersDark: boolean): ResolvedTheme {
  if (theme === Theme.Light) return 'light'
  if (theme === Theme.Dark) return 'dark'
  return systemPrefersDark ? 'dark' : 'light'
}

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
      const resolved = resolve(theme, media.matches)
      document.documentElement.setAttribute('data-theme', resolved)
      // `color-scheme` é o que faz o Chromium desenhar as barras de rolagem, os
      // campos e os menus nativos na cor certa. Sem isto o papel fica escuro e
      // a barra de rolagem ao lado dele continua branca.
      document.documentElement.style.colorScheme = resolved
    }

    apply()
    media.addEventListener('change', apply)
    return () => media.removeEventListener('change', apply)
    // `theme` é lido dentro de `apply`, então o efeito precisa dele: sem a
    // dependência, trocar de claro para escuro reaplicaria a escolha antiga.
  }, [theme])

  return resolve(theme, window.matchMedia(DARK_QUERY).matches)
}
