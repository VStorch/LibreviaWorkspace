const IPV4_LOOPBACK = /^127(\.\d{1,3}){3}$/
const LOOPBACK_NAMES: ReadonlySet<string> = new Set(['localhost', '[::1]'])

/** By exact host name: `localhost.example.com` and a `localhost` in the path are remote. */
export function isLoopback(url: string): boolean {
  let hostname: string
  try {
    hostname = new URL(url).hostname.toLowerCase()
  } catch {
    return false
  }
  return LOOPBACK_NAMES.has(hostname) || IPV4_LOOPBACK.test(hostname)
}
