import type { AiModel } from '@services/ai/provider.js'

export class AiModelCache {
  private readonly models = new Map<string, readonly AiModel[]>()

  async get(providerId: string, load: () => Promise<readonly AiModel[]>): Promise<readonly AiModel[]> {
    return this.models.get(providerId) ?? this.remember(providerId, await load())
  }

  remember(providerId: string, models: readonly AiModel[]): readonly AiModel[] {
    if (models.length > 0) this.models.set(providerId, models)
    return models
  }

  forget(providerId: string | undefined): void {
    if (providerId !== undefined) this.models.delete(providerId)
  }
}
