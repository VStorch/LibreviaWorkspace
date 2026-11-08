import { create } from 'zustand'
import type { RecentFile } from '@shared/types.js'

export const useRecents = create<{ recents: readonly RecentFile[] }>(() => ({ recents: [] }))

export function setRecents(recents: readonly RecentFile[]): void {
  useRecents.setState({ recents })
}

/** Without a warning: a list that failed to refresh only stays as it was. */
export async function refreshRecents(): Promise<void> {
  const result = await window.api.recent.list({}).catch(() => null)
  if (result?.ok === true) setRecents(result.data.files)
}
