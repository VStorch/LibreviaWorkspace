/** Nome exibido do aplicativo. */
export const APP_NAME = 'Librevia'

export const WINDOW_DEFAULTS = {
  width: 1280,
  height: 860,
  minWidth: 940,
  minHeight: 600,
} as const

/** Esquemas de URL que podem ser abertos no navegador do sistema. */
export const ALLOWED_EXTERNAL_PROTOCOLS = ['http:', 'https:', 'mailto:'] as const
