import { BrowserWindow } from 'electron'
import type { MenuCommandPayload } from '@shared/api.js'
import { MAX_OPEN_TABS } from '@services/tabs/tab-order.js'
import { attachMcpWindow } from './ai/mcp.js'
import { startWindowWith } from './external-files.js'
import { watchWindowFailures } from './failures.js'
import { createMainWindow } from './window.js'

/** As many as a window holds tabs: past that, a runaway page could not fill the screen with windows. */
export const MAX_WINDOWS = MAX_OPEN_TABS

/** The local MCP server answers through the window in front, which shows the active tab. */
export function openAppWindow(first?: MenuCommandPayload): BrowserWindow | null {
  if (BrowserWindow.getAllWindows().length >= MAX_WINDOWS) return null
  const window = createMainWindow()
  watchWindowFailures(window)
  if (first !== undefined) startWindowWith(window, first)
  window.on('focus', () => attachMcpWindow(window.webContents))
  attachMcpWindow(window.webContents)
  return window
}
