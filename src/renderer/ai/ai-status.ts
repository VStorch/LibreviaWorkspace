import type { MessageKey } from '@shared/i18n/index.js'
import type { AiPhase } from '../document/useAiRequest.js'

/** What a request shows while it runs or after it stops without an answer. */
export const AI_STATUS_LABELS: Partial<Record<AiPhase, MessageKey>> = {
  running: 'ai.panel.running',
  canceled: 'ai.panel.canceled',
  declined: 'ai.panel.declined',
  failed: 'ai.panel.failed',
}
