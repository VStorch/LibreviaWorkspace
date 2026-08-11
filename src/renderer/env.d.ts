import type { AppApi } from '@shared/api.js'

declare global {
  interface Window {
    readonly api: AppApi
  }
}

export {}
