import { describe, expect, it } from 'vitest'
import { takeLines } from './lines.js'
import { MAX_AI_TOOL_ARGUMENTS_LENGTH, MAX_AI_TOOL_CALLS } from '@shared/limits.js'
import { OpenAiToolCalls, parseOpenAiLine } from './sse.js'

const chunk = (content: string): string => `data: ${JSON.stringify({ choices: [{ delta: { content } }] })}`

describe('parseOpenAiLine', () => {
  it('reads the text chunk of each event', () => {
    expect(parseOpenAiLine(chunk('Olá'))).toEqual({ type: 'text', text: 'Olá' })
  })

  it('recognizes the end', () => {
    expect(parseOpenAiLine('data: [DONE]')).toEqual({ type: 'done' })
  })

  it('returns the error the server sends mid-stream', () => {
    expect(parseOpenAiLine('data: {"error":{"message":"model not loaded"}}')).toEqual({
      type: 'error',
      message: 'model not loaded',
    })
  })

  it('the server content filter becomes a refusal', () => {
    expect(parseOpenAiLine('data: {"choices":[{"delta":{},"finish_reason":"content_filter"}]}')).toEqual({
      type: 'refused',
    })
  })

  it('ignores comments, events without text and broken JSON', () => {
    expect(parseOpenAiLine(': keep-alive')).toEqual({ type: 'skip' })
    expect(parseOpenAiLine('data: {"choices":[{"delta":{"role":"assistant"}}]}')).toEqual({ type: 'skip' })
    expect(parseOpenAiLine('data: {"choi')).toEqual({ type: 'skip' })
  })
})

describe('takeLines', () => {
  it('keeps the split line until the next chunk', () => {
    const first = takeLines(`${chunk('Um')}\n\ndata: {"choi`)
    expect(first.lines).toEqual([chunk('Um')])

    const second = takeLines(`${first.rest}ces":[{"delta":{"content":" dois"}}]}\r\n`)
    expect(second.lines.map(parseOpenAiLine)).toEqual([{ type: 'text', text: ' dois' }])
    expect(second.rest).toBe('')
  })
})

const toolChunk = (calls: unknown[]): string =>
  `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls } }] })}`

describe('tool calls in the OpenAI stream', () => {
  it('joins the arguments pieces by index, with id and name from the first piece', () => {
    const calls = new OpenAiToolCalls()
    const lines = [
      toolChunk([
        { index: 0, id: 'a', type: 'function', function: { name: 'search_document', arguments: '' } },
      ]),
      toolChunk([{ index: 1, id: 'b', function: { name: 'read_paragraphs', arguments: '{"from":' } }]),
      toolChunk([{ index: 0, function: { arguments: '{"query":"gato"}' } }]),
      toolChunk([{ index: 1, function: { arguments: '1,"count":2}' } }]),
    ]
    for (const line of lines) {
      const event = parseOpenAiLine(line)
      if (event.type === 'tools') calls.add(event.deltas)
    }
    expect(calls.finished()).toEqual([
      { id: 'a', name: 'search_document', arguments: '{"query":"gato"}' },
      { id: 'b', name: 'read_paragraphs', arguments: '{"from":1,"count":2}' },
    ])
  })

  it('gives an id to a call that came without one, different in each round', () => {
    const first = new OpenAiToolCalls(0)
    const second = new OpenAiToolCalls(1)
    first.add([{ index: 0, name: 'x', arguments: '{}' }])
    second.add([{ index: 0, name: 'x', arguments: '{}' }])
    expect(first.finished()).toEqual([{ id: 'call_0_0', name: 'x', arguments: '{}' }])
    expect(second.finished()[0]?.id).toBe('call_1_0')
  })

  it('refuses calls beyond the limit of one response', () => {
    const calls = new OpenAiToolCalls()
    const deltas = Array.from({ length: MAX_AI_TOOL_CALLS }, (_, index) => ({ index, name: 'x' }))
    expect(calls.add(deltas)).toBe(true)
    expect(calls.add([{ index: MAX_AI_TOOL_CALLS, name: 'x' }])).toBe(false)
  })

  it('refuses arguments longer than the limit', () => {
    const calls = new OpenAiToolCalls()
    expect(calls.add([{ index: 0, name: 'x', arguments: 'a'.repeat(MAX_AI_TOOL_ARGUMENTS_LENGTH) }])).toBe(
      true,
    )
    expect(calls.add([{ index: 0, arguments: 'a' }])).toBe(false)
  })

  it('text and call in the same chunk arrive together', () => {
    const line = `data: ${JSON.stringify({
      choices: [
        { delta: { content: 'Vou ver.', tool_calls: [{ index: 0, id: 'a', function: { name: 'x' } }] } },
      ],
    })}`
    expect(parseOpenAiLine(line)).toEqual({
      type: 'tools',
      text: 'Vou ver.',
      deltas: [{ index: 0, id: 'a', name: 'x' }],
    })
  })
})
