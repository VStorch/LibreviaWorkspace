import { Component, type ReactNode } from 'react'
import { t } from '../i18n.js'

interface FailureBoundaryState {
  readonly failed: boolean
}

export class FailureBoundary extends Component<{ readonly children: ReactNode }, FailureBoundaryState> {
  override state: FailureBoundaryState = { failed: false }

  static getDerivedStateFromError(): FailureBoundaryState {
    return { failed: true }
  }

  override render(): ReactNode {
    if (!this.state.failed) return this.props.children
    return (
      <div className="banner" role="alert">
        <div className="banner__text">
          <strong>{t('shell.failure.screen.title')}</strong>
          <span className="banner__detail">{t('shell.failure.screen.detail')}</span>
        </div>
        <div className="banner__actions">
          <button type="button" className="banner__action" onClick={() => window.location.reload()}>
            {t('shell.failure.window.reload')}
          </button>
        </div>
      </div>
    )
  }
}
