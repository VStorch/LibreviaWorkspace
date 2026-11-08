/**
 * Runs `build` on every call, but hands to `install` only the result of the newest call: an older
 * one that finishes late read an older state.
 */
export function latestOnly<T>(build: () => Promise<T>, install: (result: T) => void): () => Promise<void> {
  let latest = 0
  return async () => {
    latest += 1
    const call = latest
    const result = await build()
    if (call === latest) install(result)
  }
}
