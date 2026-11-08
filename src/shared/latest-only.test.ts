import { describe, expect, it } from 'vitest'
import { latestOnly } from './latest-only.js'

describe('latestOnly', () => {
  it('installs only the newest result when an older call finishes last', async () => {
    const pending: ((value: string) => void)[] = []
    const installed: string[] = []
    const refresh = latestOnly(
      () => new Promise<string>((resolve) => pending.push(resolve)),
      (result) => installed.push(result),
    )

    const older = refresh()
    const newer = refresh()
    pending[1]!('new')
    await newer
    pending[0]!('old')
    await older

    expect(installed).toEqual(['new'])
  })

  it('installs every result when the calls do not overlap', async () => {
    const installed: number[] = []
    let count = 0
    const refresh = latestOnly(
      async () => ++count,
      (result) => installed.push(result),
    )

    await refresh()
    await refresh()

    expect(installed).toEqual([1, 2])
  })
})
