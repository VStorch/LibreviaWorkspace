import { APP_NAME } from '@shared/constants.js'
import type { RecentFile } from '@shared/types.js'
import { Icon, type IconName } from '../components/Icon.js'
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

function Tile({
  icon,
  title,
  hint,
  onClick,
}: {
  readonly icon: IconName
  readonly title: string
  readonly hint: string
  readonly onClick: () => void
}): React.JSX.Element {
  return (
    <button type="button" className="tile" onClick={onClick}>
      <span className="tile__icon">
        <Icon name={icon} />
      </span>
      <span className="tile__title">{title}</span>
      <span className="tile__hint">{hint}</span>
    </button>
  )
}

export function HomePage(): React.JSX.Element {
  const { recents, newDocument, newSpreadsheet, openViaDialog, clearRecents } = useWorkspace()

  return (
    <div className="home">
      <header className="home__header">
        <h1>{APP_NAME}</h1>
        <p className="home__subtitle">Documentos e planilhas, sem depender de nuvem.</p>
      </header>

      <div className="home__actions">
        <Tile icon="file-document" title="Novo documento" hint="Ctrl+N" onClick={() => void newDocument()} />
        <Tile
          icon="file-spreadsheet"
          title="Nova planilha"
          hint="Ctrl+Shift+N"
          onClick={() => void newSpreadsheet()}
        />
        <Tile icon="folder-open" title="Abrir arquivo" hint="Ctrl+O" onClick={() => void openViaDialog()} />
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
