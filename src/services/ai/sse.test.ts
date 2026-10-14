import { describe, expect, it } from 'vitest'
import { takeLines } from './lines.js'
import { parseOpenAiLine } from './sse.js'

const chunk = (content: string): string => `data: ${JSON.stringify({ choices: [{ delta: { content } }] })}`

describe('parseOpenAiLine', () => {
  it('lê o trecho de texto de cada evento', () => {
    expect(parseOpenAiLine(chunk('Olá'))).toEqual({ type: 'text', text: 'Olá' })
  })

  it('reconhece o fim', () => {
    expect(parseOpenAiLine('data: [DONE]')).toEqual({ type: 'done' })
  })

  it('devolve o erro que o servidor manda no meio do fluxo', () => {
    expect(parseOpenAiLine('data: {"error":{"message":"modelo não carregado"}}')).toEqual({
      type: 'error',
      message: 'modelo não carregado',
    })
  })

  it('filtro de conteúdo do servidor vira recusa', () => {
    expect(parseOpenAiLine('data: {"choices":[{"delta":{},"finish_reason":"content_filter"}]}')).toEqual({
      type: 'refused',
    })
  })

  it('ignora comentários, eventos sem texto e JSON quebrado', () => {
    expect(parseOpenAiLine(': keep-alive')).toEqual({ type: 'skip' })
    expect(parseOpenAiLine('data: {"choices":[{"delta":{"role":"assistant"}}]}')).toEqual({ type: 'skip' })
    expect(parseOpenAiLine('data: {"choi')).toEqual({ type: 'skip' })
  })
})

describe('takeLines', () => {
  it('guarda a linha partida até o próximo pedaço', () => {
    const first = takeLines(`${chunk('Um')}\n\ndata: {"choi`)
    expect(first.lines).toEqual([chunk('Um')])

    const second = takeLines(`${first.rest}ces":[{"delta":{"content":" dois"}}]}\r\n`)
    expect(second.lines.map(parseOpenAiLine)).toEqual([{ type: 'text', text: ' dois' }])
    expect(second.rest).toBe('')
  })
})
