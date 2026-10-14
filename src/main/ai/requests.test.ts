import { describe, expect, it } from 'vitest'
import { AiRequests } from './requests.js'

const REQUEST_ID = '0b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f'
const owner = (id: number) => ({ id, once: () => undefined })

describe('AiRequests', () => {
  it('só a janela que fez o pedido pode cancelá-lo', () => {
    const requests = new AiRequests()
    const signal = requests.begin(REQUEST_ID, owner(1))
    expect(requests.cancel(REQUEST_ID, 2)).toBe(false)
    expect(signal.aborted).toBe(false)
    expect(requests.cancel(REQUEST_ID, 1)).toBe(true)
    expect(signal.aborted).toBe(true)
  })
})
