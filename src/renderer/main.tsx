import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { translate, Language } from '@shared/i18n/index.js'
import { App } from './app/App.js'
import { FailureBoundary } from './components/FailureBoundary.js'
import './tokens.css'
import './styles.css'

const container = document.getElementById('root')
if (container === null) throw new Error(translate(Language.Portuguese, 'shell.error.rootNotFound'))

createRoot(container).render(
  <StrictMode>
    <FailureBoundary>
      <App />
    </FailureBoundary>
  </StrictMode>,
)
