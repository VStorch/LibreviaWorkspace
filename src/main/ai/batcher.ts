export interface DeltaBatcher {
  add(text: string): void
  flush(): void
}

/** One IPC message per token would cost more than the text; ~50 ms still reads as live. */
export function createDeltaBatcher(send: (text: string) => void, intervalMs: number): DeltaBatcher {
  let pending = ''
  let timer: ReturnType<typeof setTimeout> | null = null
  const flush = (): void => {
    if (timer !== null) clearTimeout(timer)
    timer = null
    if (pending === '') return
    const text = pending
    pending = ''
    send(text)
  }
  return {
    add(text) {
      pending += text
      timer ??= setTimeout(flush, intervalMs)
    },
    flush,
  }
}
