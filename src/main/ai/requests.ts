import { AppError, ErrorCode } from '@shared/errors.js'
import { t } from '../i18n.js'

export interface RequestOwner {
  readonly id: number
  once(event: 'destroyed', listener: () => void): unknown
}

/** A closed window cancels what it asked for: nobody would see the answer, and it may cost money. */
export class AiRequests {
  private readonly active = new Map<string, { controller: AbortController; owner: number }>()
  private readonly watched = new Set<number>()

  begin(requestId: string, owner: RequestOwner): AbortSignal {
    if (this.active.has(requestId)) throw new AppError(ErrorCode.InvalidRequest, t('errors.ipc.invalidData'))
    const controller = new AbortController()
    this.active.set(requestId, { controller, owner: owner.id })
    if (!this.watched.has(owner.id)) {
      this.watched.add(owner.id)
      owner.once('destroyed', () => {
        this.watched.delete(owner.id)
        this.cancelOwnedBy(owner.id)
      })
    }
    return controller.signal
  }

  cancel(requestId: string, owner: number): boolean {
    const request = this.active.get(requestId)
    if (request?.owner !== owner) return false
    request.controller.abort()
    return true
  }

  end(requestId: string): void {
    this.active.delete(requestId)
  }

  private cancelOwnedBy(owner: number): void {
    for (const [id, request] of this.active) {
      if (request.owner === owner) this.cancel(id, owner)
    }
  }
}
