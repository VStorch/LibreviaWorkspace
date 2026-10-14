/** Where the focus goes for a key pressed inside a menu of `count` items; `null` for any other key. */
export function nextMenuIndex(current: number, count: number, key: string): number | null {
  if (count === 0) return null
  switch (key) {
    case 'ArrowDown':
      return current < 0 ? 0 : (current + 1) % count
    case 'ArrowUp':
      return current < 0 ? count - 1 : (current - 1 + count) % count
    case 'Home':
      return 0
    case 'End':
      return count - 1
    default:
      return null
  }
}
