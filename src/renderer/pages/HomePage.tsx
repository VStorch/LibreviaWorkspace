import { useState } from 'react'
import { APP_NAME } from '@shared/constants.js'
import swallow from '../../../build/icon.svg?raw'
import type { RecentFile } from '@shared/types.js'
import { Icon, type IconName } from '../components/Icon.js'
import { useT } from '../i18n.js'
import { useRecents } from '../state/recents.js'
import {
  clearRecents,
  newDocument,
  newSpreadsheet,
  openRecent,
  openViaDialog,
  setTemplateGallery,
} from '../state/shell.js'

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
  return (
    <li>
      <button type="button" className="recent" onClick={() => void openRecent(file.path)}>
        <span className="recent__name">{file.name}</span>
        {/* The path tells apart files with the same name on network shares. */}
        <span className="recent__path" title={file.path}>
          {file.path}
        </span>
        <span className="recent__when">{formatWhen(file.openedAt)}</span>
      </button>
    </li>
  )
}

const CLICKS_TO_FLAP = 3

/** The app icon itself, so the home screen and the system show the same swallow. */
function BrandMark(): React.JSX.Element {
  const [flapping, setFlapping] = useState(false)
  return (
    <span
      className={flapping ? 'home__mark home__mark--flapping' : 'home__mark'}
      aria-hidden="true"
      // Repeated clicks would otherwise select the nearest text, the title.
      onMouseDown={(event) => event.preventDefault()}
      onClick={(event) => {
        if (event.detail === CLICKS_TO_FLAP) setFlapping(true)
      }}
      onAnimationEnd={() => setFlapping(false)}
      dangerouslySetInnerHTML={{ __html: swallow }}
    />
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
  const t = useT()
  const recents = useRecents((state) => state.recents)

  return (
    <div className="home">
      <header className="home__header">
        <div className="home__brand">
          <BrandMark />
          <h1>{APP_NAME}</h1>
        </div>
        <p className="home__subtitle">{t('shell.home.subtitle')}</p>
      </header>

      <div className="home__actions">
        <Tile
          icon="file-document"
          tone="document"
          title={t('shell.home.newDocument')}
          hint="Ctrl+N"
          onClick={newDocument}
        />
        <Tile
          icon="file-spreadsheet"
          tone="spreadsheet"
          title={t('shell.home.newSpreadsheet')}
          hint="Ctrl+Shift+N"
          onClick={newSpreadsheet}
        />
        <Tile
          icon="folder-open"
          tone="neutral"
          title={t('shell.home.openFile')}
          hint="Ctrl+O"
          onClick={() => void openViaDialog()}
        />
        <Tile
          icon="file-document"
          tone="neutral"
          title={t('shell.home.fromTemplate')}
          hint=""
          onClick={() => setTemplateGallery(true)}
        />
      </div>

      <section className="home__recents">
        <div className="home__recents-head">
          <h2>{t('shell.home.recentFiles')}</h2>
          {recents.length > 0 && (
            <button type="button" className="link" onClick={() => void clearRecents()}>
              {t('shell.home.clear')}
            </button>
          )}
        </div>

        {recents.length === 0 ? (
          <p className="home__empty">{t('shell.home.emptyRecents')}</p>
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
