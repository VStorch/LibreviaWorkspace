import { create } from 'zustand'

export const useOptionsDialog = create<{ readonly open: boolean }>(() => ({ open: false }))

export function setOptionsDialog(open: boolean): void {
  useOptionsDialog.setState({ open })
}
