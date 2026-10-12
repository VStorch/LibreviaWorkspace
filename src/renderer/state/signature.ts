import { anonymousName } from './privacy-actions.js'
import { currentPreferences } from './preferences.js'
import { useWorkspace } from './workspace.js'

/**
 * The name new comments and tracked changes carry. Generic in a file that asks not to keep personal
 * information, as Word does; empty when the person never gave one.
 */
export function signingName(): string {
  if (useWorkspace.getState().properties?.removePersonalInformation === true) return anonymousName()
  return currentPreferences().authorName.trim()
}
