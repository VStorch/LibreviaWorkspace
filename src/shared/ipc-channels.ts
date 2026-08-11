/**
 * Nomes de canal IPC.
 *
 * Este arquivo é deliberadamente livre de dependências: ele é importado pelo
 * preload, que roda em contexto sandboxed e não pode carregar pacotes de
 * terceiros. Os schemas de validação ficam em `ipc.ts`, que só o main importa.
 */
export const IpcChannel = {
  /** Canal de fumaça da Fase 0: valida ida e volta e o envelope de resultado. */
  AppPing: 'app:ping',
} as const

export type IpcChannel = (typeof IpcChannel)[keyof typeof IpcChannel]

export const ALL_IPC_CHANNELS: readonly IpcChannel[] = Object.values(IpcChannel)
