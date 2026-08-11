import { APP_NAME } from '@shared/constants.js'
import type { RecentFile } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'

function formatWhen(timestamp: number): string {
  return new Date(timestamp).toLocaleString('pt-BR', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

function RecentItem({ file }: { readonly file: RecentFile }): React.JSX.Element {
  const openRecent = useWorkspace((state) => state.openRecent)

  return (
    <li>
      <button type="button" className="recent" onClick={() => void openRecent(file.path)}>
        <span className="recent__name">{file.name}</span>
        {/* O caminho é o que distingue dois arquivos de mesmo nome em pastas
            diferentes — situação comum em unidade de rede compartilhada. */}
        <span className="recent__path" title={file.path}>
          {file.path}
        </span>
        <span className="recent__when">{formatWhen(file.openedAt)}</span>
      </button>
    </li>
  )
}

export function HomePage(): React.JSX.Element {
  const { recents, newDocument, openViaDialog, clearRecents } = useWorkspace()

  return (
    <div className="home">
      <header className="home__header">
        <h1>{APP_NAME}</h1>
        <p className="home__subtitle">Documentos e planilhas, sem depender de nuvem.</p>
      </header>

      <div className="home__actions">
        <button type="button" className="tile" onClick={() => void newDocument()}>
          <span className="tile__title">Novo documento</span>
          <span className="tile__hint">Ctrl+N</span>
        </button>

        <button type="button" className="tile" disabled title="Disponível na Fase 5">
          <span className="tile__title">Nova planilha</span>
          <span className="tile__hint">em breve</span>
        </button>

        <button type="button" className="tile" onClick={() => void openViaDialog()}>
          <span className="tile__title">Abrir arquivo</span>
          <span className="tile__hint">Ctrl+O</span>
        </button>
      </div>

      <section className="home__recents">
        <div className="home__recents-head">
          <h2>Arquivos recentes</h2>
          {recents.length > 0 && (
            <button type="button" className="link" onClick={() => void clearRecents()}>
              Limpar
            </button>
          )}
        </div>

        {recents.length === 0 ? (
          <p className="home__empty">Nenhum arquivo aberto ainda.</p>
        ) : (
          <ul className="home__recents-list">
            {recents.map((file) => (
              <RecentItem key={file.path} file={file} />
            ))}
          </ul>
        )}
      </section>
    </div>
  )
}
