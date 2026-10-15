import { constants } from 'node:fs'
import { access as fsAccess } from 'node:fs/promises'
import { delimiter, join } from 'node:path'
import { ACP_AGENTS, type AcpAgentId } from '@services/ai/acp-agents.js'
import { joinArgs } from '@services/ai/acp-session.js'
import { AiProviderKind } from '@services/ai/provider.js'

export const CLAUDE_COMMAND = 'claude'
const PROBE_TIMEOUT_MS = 800

const LOCAL_SERVERS = [
  { label: 'Ollama', baseUrl: 'http://localhost:11434/v1' },
  { label: 'LM Studio', baseUrl: 'http://localhost:1234/v1' },
] as const

export interface AiCandidate {
  readonly kind: AiProviderKind
  readonly label: string
  readonly baseUrl?: string
  readonly agentId?: AcpAgentId
  readonly args?: string
}

export interface DiscoveryDeps {
  readonly fetch: typeof fetch
  readonly findExecutable: (name: string) => Promise<string | null>
}

type Access = (path: string, mode: number) => Promise<void>

export async function findExecutable(
  name: string,
  env: NodeJS.ProcessEnv = process.env,
  access: Access = fsAccess,
): Promise<string | null> {
  const names = process.platform === 'win32' ? [`${name}.exe`] : [name]
  const folders = (env['PATH'] ?? '').split(delimiter).filter((folder) => folder !== '')
  for (const folder of folders) {
    for (const file of names) {
      const path = join(folder, file)
      if (
        await access(path, constants.X_OK).then(
          () => true,
          () => false,
        )
      )
        return path
    }
  }
  return null
}

export async function discoverProviders(
  deps: DiscoveryDeps = { fetch, findExecutable: (name) => findExecutable(name) },
): Promise<AiCandidate[]> {
  const servers = LOCAL_SERVERS.map(async (server): Promise<AiCandidate[]> =>
    (await answers(deps.fetch, server.baseUrl))
      ? [{ kind: AiProviderKind.OpenAiCompatible, label: server.label, baseUrl: server.baseUrl }]
      : [],
  )
  const cli = deps
    .findExecutable(CLAUDE_COMMAND)
    .then((path): AiCandidate[] =>
      path === null ? [] : [{ kind: AiProviderKind.ClaudeCli, label: 'Claude Code' }],
    )
  const agents = ACP_AGENTS.map(async (agent): Promise<AiCandidate[]> =>
    (await deps.findExecutable(agent.command)) === null
      ? []
      : [{ kind: AiProviderKind.Acp, label: agent.label, agentId: agent.id, args: joinArgs(agent.args) }],
  )
  return (await Promise.all([...servers, cli, ...agents])).flat()
}

async function answers(fetcher: typeof fetch, baseUrl: string): Promise<boolean> {
  try {
    const response = await fetcher(`${baseUrl}/models`, {
      signal: AbortSignal.timeout(PROBE_TIMEOUT_MS),
      redirect: 'error',
    })
    await response.body?.cancel()
    return response.ok
  } catch {
    return false
  }
}
