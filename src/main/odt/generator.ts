import { app } from 'electron'
import { APP_NAME } from '@shared/constants.js'

/** `meta:generator`, as LibreOffice writes its own: the program and its version. */
export function odtGenerator(): string {
  return `${APP_NAME}/${app.getVersion()}`
}
