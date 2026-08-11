import { useEffect, useState } from 'react'
import { runSecurityProbes, type ProbeResult } from './security-probe.js'

interface PingState {
  readonly status: 'carregando' | 'ok' | 'erro'
  readonly detail: string
  readonly versions?: Record<string, string>
}

export function App(): React.JSX.Element {
  const [probes, setProbes] = useState<readonly ProbeResult[]>([])
  const [ping, setPing] = useState<PingState>({ status: 'carregando', detail: 'chamando o processo main…' })

  useEffect(() => {
    setProbes(runSecurityProbes())

    void window.api.app
      .ping({ message: 'fase 0' })
      .then((result) => {
        setPing(
          result.ok
            ? {
                status: 'ok',
                detail: `resposta “${result.data.echo}” recebida do processo main`,
                versions: result.data.versions,
              }
            : { status: 'erro', detail: result.error.message },
        )
      })
      .catch(() => {
        setPing({ status: 'erro', detail: 'A ponte IPC não respondeu.' })
      })
  }, [])

  const allProbesPassed = probes.length > 0 && probes.every((probe) => probe.passed)

  return (
    <main className="shell">
      <header className="shell__header">
        <h1>Librevia</h1>
        <p className="shell__subtitle">Fase 0 — fundação. Nome provisório.</p>
      </header>

      <section className="card">
        <h2>Isolamento do renderer</h2>
        <p className="card__lead">
          {allProbesPassed
            ? 'O renderer não alcança o Node.js.'
            : 'Verificação pendente ou falhando — ver detalhes abaixo.'}
        </p>
        <ul className="checks">
          {probes.map((probe) => (
            <li key={probe.label} className={probe.passed ? 'checks__item--ok' : 'checks__item--fail'}>
              <span className="checks__mark">{probe.passed ? '✓' : '✕'}</span>
              <code>{probe.label}</code>
              <span className="checks__note">{probe.expectation}</span>
            </li>
          ))}
        </ul>
      </section>

      <section className="card">
        <h2>Ponte IPC</h2>
        <p className={`card__lead card__lead--${ping.status}`}>{ping.detail}</p>
        {ping.versions !== undefined && (
          <dl className="versions">
            {Object.entries(ping.versions).map(([name, value]) => (
              <div key={name} className="versions__row">
                <dt>{name}</dt>
                <dd>{value}</dd>
              </div>
            ))}
          </dl>
        )}
      </section>

      <footer className="shell__footer">
        Próxima etapa: Fase 1 — menu nativo, tela inicial e sistema de arquivos.
      </footer>
    </main>
  )
}
