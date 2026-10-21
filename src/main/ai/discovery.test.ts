import { constants } from 'node:fs'
import { delimiter, join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { discoverProviders, findExecutable } from './discovery.js'

describe('discoverProviders', () => {
  it('finds the server that answers and the program in PATH', async () => {
    const asked: string[] = []
    const candidates = await discoverProviders({
      fetch: (async (url: string) => {
        asked.push(url)
        if (url.includes('11434')) return new Response('{}')
        throw new TypeError('connection refused')
      }) as unknown as typeof fetch,
      findExecutable: async (name) => (name === 'claude' || name === 'gemini' ? `/usr/bin/${name}` : null),
    })
    expect(asked.sort()).toEqual(['http://localhost:11434/v1/models', 'http://localhost:1234/v1/models'])
    expect(candidates).toEqual([
      { kind: 'openai-compatible', label: 'Ollama', baseUrl: 'http://localhost:11434/v1' },
      { kind: 'claude-cli', label: 'Claude Code' },
      { kind: 'acp', label: 'Gemini CLI', agentId: 'gemini', args: '--acp' },
    ])
  })

  it('looks up each known ACP agent by program name, never by npx', async () => {
    const searched: string[] = []
    await discoverProviders({
      fetch: (async () => {
        throw new TypeError('connection refused')
      }) as unknown as typeof fetch,
      findExecutable: async (name) => {
        searched.push(name)
        return null
      },
    })
    expect(searched).toContain('codex-acp')
    expect(searched).toContain('opencode')
    expect(searched).not.toContain('npx')
    expect(searched).not.toContain('uvx')
  })

  it('gives up on a server that does not answer in time', async () => {
    const started = Date.now()
    const candidates = await discoverProviders({
      fetch: ((_url: string, init: RequestInit) =>
        new Promise((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => reject(new Error('timeout')))
        })) as unknown as typeof fetch,
      findExecutable: async () => null,
    })
    expect(candidates).toEqual([])
    expect(Date.now() - started).toBeLessThan(3_000)
  })
})

describe('findExecutable', () => {
  it('only accepts an executable file in a PATH folder', async () => {
    const checked: [string, number][] = []
    const path = await findExecutable(
      'claude',
      { PATH: ['/a', '/b'].join(delimiter) },
      async (file, mode) => {
        checked.push([file, mode])
        if (file !== join('/b', 'claude')) throw new Error('ENOENT')
      },
    )
    expect(path).toBe(join('/b', 'claude'))
    expect(checked.every(([, mode]) => mode === constants.X_OK)).toBe(true)
  })
})
