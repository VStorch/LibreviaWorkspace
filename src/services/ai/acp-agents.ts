export interface AcpAgent {
  readonly id: string
  readonly label: string
  readonly command: string
  readonly args: readonly string[]
}

/**
 * From the official ACP registry, only agents that install as a program of their own: `npx` and
 * `uvx` would download the agent from the internet on every run.
 */
export const ACP_AGENTS = [
  { id: 'gemini', label: 'Gemini CLI', command: 'gemini', args: ['--acp'] },
  { id: 'claude-agent-acp', label: 'Claude Agent', command: 'claude-agent-acp', args: [] },
  { id: 'codex-acp', label: 'Codex', command: 'codex-acp', args: [] },
  { id: 'opencode', label: 'OpenCode', command: 'opencode', args: ['acp'] },
  { id: 'goose', label: 'Goose', command: 'goose', args: ['acp'] },
  { id: 'copilot', label: 'GitHub Copilot', command: 'copilot', args: ['--acp'] },
  // The registry adds `--experimental-skills`, which turns on more than the protocol.
  { id: 'qwen', label: 'Qwen Code', command: 'qwen', args: ['--acp'] },
  { id: 'kimi', label: 'Kimi CLI', command: 'kimi', args: ['acp'] },
  { id: 'cursor-agent', label: 'Cursor Agent', command: 'cursor-agent', args: ['acp'] },
  { id: 'kilo', label: 'Kilo Code', command: 'kilo', args: ['acp'] },
  { id: 'auggie', label: 'Auggie', command: 'auggie', args: ['--acp'] },
  { id: 'cline', label: 'Cline', command: 'cline', args: ['--acp'] },
  { id: 'vibe-acp', label: 'Mistral Vibe', command: 'vibe-acp', args: [] },
  { id: 'amp-acp', label: 'Amp', command: 'amp-acp', args: [] },
] as const satisfies readonly AcpAgent[]

export type AcpAgentId = (typeof ACP_AGENTS)[number]['id']

export const ACP_AGENT_IDS = ACP_AGENTS.map((agent) => agent.id) as [AcpAgentId, ...AcpAgentId[]]

export function acpAgentOf(id: AcpAgentId): AcpAgent {
  return ACP_AGENTS.find((agent) => agent.id === id) ?? ACP_AGENTS[0]
}
