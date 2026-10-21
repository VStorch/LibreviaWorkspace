import { describe, expect, it } from 'vitest'
import { parseClaudeCliLine } from './claude-stream.js'

const delta = (text: string): string =>
  JSON.stringify({
    type: 'stream_event',
    event: { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text } },
  })

describe('parseClaudeCliLine', () => {
  it('reads the partial text', () => {
    expect(parseClaudeCliLine(delta('Olá'))).toEqual({ type: 'text', text: 'Olá' })
  })

  it('reads the result with the cost', () => {
    const line = JSON.stringify({
      type: 'result',
      subtype: 'success',
      is_error: false,
      stop_reason: 'end_turn',
      total_cost_usd: 0.0123,
      result: 'Olá',
    })
    expect(parseClaudeCliLine(line)).toEqual({
      type: 'result',
      failed: false,
      refused: false,
      costUsd: 0.0123,
      message: 'Olá',
    })
  })

  it('flags the refusal and the error', () => {
    const refusal = parseClaudeCliLine(
      JSON.stringify({ type: 'result', is_error: false, stop_reason: 'refusal' }),
    )
    expect(refusal).toMatchObject({ refused: true, failed: false })
    const failure = parseClaudeCliLine(
      JSON.stringify({ type: 'result', is_error: true, result: 'Not logged in' }),
    )
    expect(failure).toMatchObject({ failed: true, message: 'Not logged in', costUsd: undefined })
  })

  it('reads the model that answered, from the session start and the message start', () => {
    expect(
      parseClaudeCliLine(JSON.stringify({ type: 'system', subtype: 'init', model: 'claude-opus-5-5' })),
    ).toEqual({ type: 'model', model: 'claude-opus-5-5' })
    expect(
      parseClaudeCliLine(
        JSON.stringify({
          type: 'stream_event',
          event: { type: 'message_start', message: { model: 'claude-opus-5-5' } },
        }),
      ),
    ).toEqual({ type: 'message', model: 'claude-opus-5-5' })
  })

  it('the start of each message marks a new round, even without the model', () => {
    expect(
      parseClaudeCliLine(JSON.stringify({ type: 'stream_event', event: { type: 'message_start' } })),
    ).toEqual({ type: 'message', model: undefined })
  })

  it('ignores the other lines', () => {
    expect(parseClaudeCliLine(JSON.stringify({ type: 'system', subtype: 'init' }))).toEqual({ type: 'skip' })
    expect(parseClaudeCliLine(JSON.stringify({ type: 'assistant', message: {} }))).toEqual({ type: 'skip' })
    expect(parseClaudeCliLine('not JSON')).toEqual({ type: 'skip' })
  })
})
