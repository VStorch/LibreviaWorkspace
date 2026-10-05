import { defineConfig } from '@playwright/test'

/**
 * End-to-end tests against the real Electron app.
 *
 * They run on `out/`, not on the source: what needs proving is the packaged app, with sandboxed
 * preload, contextIsolation and the published .NET sidecar. A test running on the source would
 * cross none of those boundaries, which are exactly where this architecture's bugs show up.
 *
 * A single worker, no parallelism: each test launches a whole Electron, and two at once would fight
 * over the same `userData` and single-instance lock.
 */
export default defineConfig({
  testDir: './e2e',
  fullyParallel: false,
  workers: 1,
  // Starting Electron and the sidecar takes a few seconds; the 30 s default leaves little margin on
  // a loaded machine.
  timeout: 90_000,
  expect: { timeout: 15_000 },
  reporter: [['list']],
  // Fails on CI if someone forgets a `.only`: the rest of the suite would not run.
  forbidOnly: Boolean(process.env.CI),
})
