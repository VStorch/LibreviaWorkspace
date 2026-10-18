import { describe, expect, it } from 'vitest'
import {
  answerPermission,
  isAuthError,
  joinArgs,
  modelsOf,
  refusePermission,
  splitArgs,
  textOfUpdate,
} from './acp-session.js'
import type { AiToolServer } from './provider.js'

describe('textOfUpdate', () => {
  it('só o texto da resposta passa; pensamento e ferramentas ficam de fora', () => {
    const text = { type: 'text', text: 'olá' }
    expect(textOfUpdate({ sessionUpdate: 'agent_message_chunk', content: text })).toBe('olá')
    expect(textOfUpdate({ sessionUpdate: 'agent_thought_chunk', content: text })).toBe('')
    expect(textOfUpdate({ sessionUpdate: 'tool_call', content: [] })).toBe('')
    expect(
      textOfUpdate({ sessionUpdate: 'agent_message_chunk', content: { type: 'image', data: 'x' } }),
    ).toBe('')
  })
})

describe('refusePermission', () => {
  it('escolhe a opção de recusar uma vez quando o agente oferece', () => {
    const answer = refusePermission([
      { optionId: 'sim', kind: 'allow_once' },
      { optionId: 'nunca', kind: 'reject_always' },
      { optionId: 'nao', kind: 'reject_once' },
    ])
    expect(answer).toEqual({ outcome: { outcome: 'selected', optionId: 'nao' } })
  })

  it('sem opção de recusa, cancela o pedido em vez de permitir', () => {
    expect(refusePermission([{ optionId: 'sim', kind: 'allow_always' }])).toEqual({
      outcome: { outcome: 'cancelled' },
    })
  })
})

describe('answerPermission', () => {
  const SERVER: AiToolServer = {
    name: 'librevia',
    command: '/bin/librevia',
    args: [],
    env: {},
    tools: ['read_paragraphs', 'add_comment'],
  }
  const OPTIONS = [
    { optionId: 'sempre', kind: 'allow_always' },
    { optionId: 'sim', kind: 'allow_once' },
    { optionId: 'nao', kind: 'reject_once' },
  ]
  const allowed = { outcome: { outcome: 'selected', optionId: 'sim' } }
  const refused = { outcome: { outcome: 'selected', optionId: 'nao' } }

  it('permite uma vez as ferramentas do Librevia, pelo nome que cada agente dá', () => {
    for (const call of [
      { name: 'mcp__librevia__read_paragraphs', title: 'mcp__librevia__read_paragraphs', kind: 'other' },
      { title: 'mcp_librevia_add_comment' },
      { title: ' read_paragraphs (librevia MCP Server) ', kind: 'other' },
    ]) {
      expect(answerPermission(call, OPTIONS, SERVER)).toEqual(allowed)
    }
  })

  it('recusa o resto: nome sem o servidor, outro servidor, tipo que executa, mensagem sem as ferramentas', () => {
    for (const call of [
      { name: 'add_comment', kind: 'other' },
      { title: 'add_comment(text: "oi")' },
      { title: 'add_comment (github MCP Server)' },
      { name: 'Bash', title: 'mcp__librevia__read_paragraphs', kind: 'execute' },
      { title: 'mcp__librevia__read_paragraphs', kind: 'execute' },
      { title: 'mcp__outro__read_paragraphs' },
      { title: 'echo mcp__librevia__read_paragraphs' },
      { name: 'mcp__librevia__propose_edit' },
      {},
    ]) {
      expect(answerPermission(call, OPTIONS, SERVER)).toEqual(refused)
    }
    expect(answerPermission({ title: 'mcp__librevia__read_paragraphs' }, OPTIONS, undefined)).toEqual(refused)
  })
})

describe('modelsOf', () => {
  it('lê os modelos da opção de categoria modelo, inclusive em grupos', () => {
    const models = modelsOf([
      {
        id: 'mode',
        type: 'select',
        category: 'mode',
        currentValue: 'ask',
        options: [{ value: 'ask', name: 'Ask' }],
      },
      {
        id: 'model',
        type: 'select',
        category: 'model',
        currentValue: 'a',
        options: [{ value: 'a', name: 'Modelo A' }, { options: [{ value: 'b', name: 'Modelo B' }] }],
      },
    ])
    expect(models).toEqual([
      { id: 'a', label: 'Modelo A' },
      { id: 'b', label: 'Modelo B' },
    ])
  })

  it('sem opção de modelo, a lista fica vazia e vale o automático', () => {
    expect(modelsOf(undefined)).toEqual([])
    expect(modelsOf([{ id: 'x', type: 'boolean', category: 'model', currentValue: true }])).toEqual([])
  })
})

describe('isAuthError', () => {
  it('reconhece o código de autenticação exigida do ACP', () => {
    expect(isAuthError({ code: -32000, message: 'Authentication required' })).toBe(true)
    expect(isAuthError({ code: -32603 })).toBe(false)
    expect(isAuthError(new Error('x'))).toBe(false)
  })
})

describe('splitArgs e joinArgs', () => {
  it('separa por espaços e mantém juntos os trechos entre aspas', () => {
    expect(splitArgs(` --acp  --dir 'meus documentos' "it's" ''`)).toEqual([
      '--acp',
      '--dir',
      'meus documentos',
      "it's",
      '',
    ])
  })

  it('não interpreta nada de shell', () => {
    expect(splitArgs('$(rm -rf ~); echo `x` | cat')).toEqual([
      '$(rm',
      '-rf',
      '~);',
      'echo',
      '`x`',
      '|',
      'cat',
    ])
  })

  it('juntar e separar de novo devolve os mesmos argumentos', () => {
    const args = ['--acp', 'com espaço', "it's", '']
    expect(splitArgs(joinArgs(args))).toEqual(args)
  })
})
