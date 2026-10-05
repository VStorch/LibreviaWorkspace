import { useEffect } from 'react'
import { Theme, type ResolvedTheme } from '@shared/types.js'
import { usePreferences } from './preferences.js'

const DARK_QUERY = '(prefers-color-scheme: dark)'

/**
 * An explicit choice decides alone; **only** `system` queries the media. Letting the query always
 * decide, relying on main's `themeSource`, does not work: the end-to-end test asks for dark and
 * `data-theme` stays `light`.
 */
function resolve(theme: Theme, systemPrefersDark: boolean): ResolvedTheme {
  if (theme === Theme.Light) return 'light'
  if (theme === Theme.Dark) return 'dark'
  return systemPrefersDark ? 'dark' : 'light'
}

/** An attribute, not a class: CSS uses `:root[data-theme='dark']` in both cases. */
export function useTheme(): ResolvedTheme {
  const theme = usePreferences((state) => state.preferences.theme)

  useEffect(() => {
    const media = window.matchMedia(DARK_QUERY)

    const apply = (): void => {
      const resolved = resolve(theme, media.matches)
      document.documentElement.setAttribute('data-theme', resolved)
      // `color-scheme` puts native scrollbars, fields and menus in the right color.
      document.documentElement.style.colorScheme = resolved
    }

    apply()
    media.addEventListener('change', apply)
    return () => media.removeEventListener('change', apply)
    // Without the dependency, switching the theme would reapply the old choice.
  }, [theme])

  return resolve(theme, window.matchMedia(DARK_QUERY).matches)
}
