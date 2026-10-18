import { describe, expect, it } from 'vitest'
import { parseClaudeCliLine } from './claude-stream.js'

const delta = (text: string): string =>
  JSON.stringify({
    type: 'stream_event',
    event: { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text } },
  })

describe('parseClaudeCliLine', () => {
  it('lê o texto parcial', () => {
    expect(parseClaudeCliLine(delta('Olá'))).toEqual({ type: 'text', text: 'Olá' })
  })

  it('lê o resultado com o custo', () => {
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

  it('marca a recusa e o erro', () => {
    const refusal = parseClaudeCliLine(
      JSON.stringify({ type: 'result', is_error: false, stop_reason: 'refusal' }),
    )
    expect(refusal).toMatchObject({ refused: true, failed: false })
    const failure = parseClaudeCliLine(
      JSON.stringify({ type: 'result', is_error: true, result: 'Not logged in' }),
    )
    expect(failure).toMatchObject({ failed: true, message: 'Not logged in', costUsd: undefined })
  })

  it('lê o modelo que respondeu, do início da sessão e do início da mensagem', () => {
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

  it('o início de cada mensagem marca uma rodada nova, mesmo sem o modelo', () => {
    expect(
      parseClaudeCliLine(JSON.stringify({ type: 'stream_event', event: { type: 'message_start' } })),
    ).toEqual({ type: 'message', model: undefined })
  })

  it('ignora as outras linhas', () => {
    expect(parseClaudeCliLine(JSON.stringify({ type: 'system', subtype: 'init' }))).toEqual({ type: 'skip' })
    expect(parseClaudeCliLine(JSON.stringify({ type: 'assistant', message: {} }))).toEqual({ type: 'skip' })
    expect(parseClaudeCliLine('não é JSON')).toEqual({ type: 'skip' })
  })
})
