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

/**
 * A marca.
 *
 * Duas folhas empilhadas, uma de cada cor da suíte: a verde da planilha atrás,
 * a azul do documento à frente. É o único lugar do programa onde cabe algo
 * assim — a tela inicial é onde o Writer põe o Start Center e o Word põe a sua
 * capa. Dentro do editor, o espaço pertence ao documento.
 */
function BrandMark(): React.JSX.Element {
  return (
    <svg
      className="home__mark"
      viewBox="0 0 40 40"
      width="42"
      height="42"
      fill="none"
      strokeLinecap="round"
      aria-hidden="true"
    >
      <rect
        x="14"
        y="6"
        width="19"
        height="26"
        rx="3"
        transform="rotate(9 23.5 19)"
        fill="var(--accent-spreadsheet-soft)"
        stroke="var(--accent-spreadsheet)"
        strokeWidth="1.6"
      />
      <rect
        x="7"
        y="9"
        width="19"
        height="26"
        rx="3"
        fill="var(--surface)"
        stroke="var(--accent-document)"
        strokeWidth="1.6"
      />
      <path
        d="M11.5 16h10M11.5 21h10M11.5 26h6"
        stroke="var(--accent-document)"
        strokeWidth="1.6"
        opacity="0.45"
      />
    </svg>
  )
}

type Tone = 'document' | 'spreadsheet' | 'neutral'

function Tile({
  icon,
  tone,
  title,
  hint,
  onClick,
}: {
  readonly icon: IconName
  readonly tone: Tone
  readonly title: string
  readonly hint: string
  readonly onClick: () => void
}): React.JSX.Element {
  return (
    <button type="button" className="tile" onClick={onClick}>
      <span className={`tile__icon tile__icon--${tone}`}>
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
        <div className="home__brand">
          <BrandMark />
          <h1>{APP_NAME}</h1>
        </div>
        <p className="home__subtitle">Documentos e planilhas, sem depender de nuvem.</p>
      </header>

      <div className="home__actions">
        <Tile
          icon="file-document"
          tone="document"
          title="Novo documento"
          hint="Ctrl+N"
          onClick={() => void newDocument()}
        />
        <Tile
          icon="file-spreadsheet"
          tone="spreadsheet"
          title="Nova planilha"
          hint="Ctrl+Shift+N"
          onClick={() => void newSpreadsheet()}
        />
        <Tile
          icon="folder-open"
          tone="neutral"
          title="Abrir arquivo"
          hint="Ctrl+O"
          onClick={() => void openViaDialog()}
        />
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
