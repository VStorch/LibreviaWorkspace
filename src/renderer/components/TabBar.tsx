import { useStore } from 'zustand'
import { DocumentKind, OpenMode } from '@shared/types.js'
import { isHomeTab } from '@services/tabs/tab-order.js'
import { useT } from '../i18n.js'
import { closeTab, openHome } from '../state/shell.js'
import { activateTab, useTabs, type TabEntry } from '../state/tabs.js'
import { usePreferences } from '../state/preferences.js'
import { Icon, type IconName } from './Icon.js'

const KIND_ICON: Record<DocumentKind, IconName> = {
  [DocumentKind.Document]: 'file-document',
  [DocumentKind.Spreadsheet]: 'file-spreadsheet',
  [DocumentKind.Pdf]: 'file-pdf',
}

const MIDDLE_BUTTON = 1

/** With one file per window, a window holding a single tab needs no bar. */
export function TabBar(): React.JSX.Element | null {
  const t = useT()
  const tabs = useTabs((state) => state.tabs)
  const activeId = useTabs((state) => state.activeId)
  const oneFilePerWindow = usePreferences((state) => state.preferences.openMode === OpenMode.Windows)
  const closable = tabs.length > 1 || !isHomeTab(activeId)
  if (oneFilePerWindow && tabs.length === 1) return null

  return (
    <nav className="tabbar" aria-label={t('shell.tabs.list')}>
      <ul className="tabbar__list">
        {tabs.map((tab) => (
          <Tab key={tab.id} tab={tab} active={tab.id === activeId} closable={closable} />
        ))}
      </ul>
      <button
        type="button"
        className="tabbar__add"
        aria-label={t('shell.tabs.openHome')}
        title={t('shell.tabs.openHome')}
        onClick={openHome}
      >
        <Icon name="plus" />
      </button>
    </nav>
  )
}

function Tab({
  tab,
  active,
  closable,
}: {
  readonly tab: TabEntry
  readonly active: boolean
  readonly closable: boolean
}): React.JSX.Element {
  const t = useT()
  const home = isHomeTab(tab.id)
  const fileName = useStore(tab.store, (state) => state.file?.name ?? null)
  const kind = useStore(tab.store, (state) => state.file?.kind ?? null)
  const path = useStore(tab.store, (state) => state.file?.path ?? null)
  const dirty = useStore(tab.store, (state) => state.isDirty)
  const name = home || fileName === null ? t('shell.tabs.home') : fileName
  const icon: IconName = home || kind === null ? 'home' : KIND_ICON[kind]

  return (
    <li
      className={active ? 'tabbar__tab tabbar__tab--active' : 'tabbar__tab'}
      onMouseDown={(event) => {
        // Middle-click would otherwise start the browser's autoscroll.
        if (event.button === MIDDLE_BUTTON) event.preventDefault()
      }}
      onAuxClick={(event) => {
        if (event.button === MIDDLE_BUTTON && closable) void closeTab(tab.id)
      }}
    >
      <button
        type="button"
        className="tabbar__label"
        aria-current={active ? 'page' : undefined}
        aria-label={dirty ? t('shell.tabs.unsaved', { name }) : name}
        title={path ?? name}
        onClick={() => activateTab(tab.id)}
      >
        <Icon name={icon} />
        <span className="tabbar__name">{name}</span>
        {dirty && (
          <span className="tabbar__dirty" aria-hidden="true">
            •
          </span>
        )}
      </button>
      {closable && (
        <button
          type="button"
          className="tabbar__close"
          aria-label={t('shell.tabs.close', { name })}
          title={t('shell.tabs.close', { name })}
          onClick={() => void closeTab(tab.id)}
        >
          <Icon name="close" />
        </button>
      )}
    </li>
  )
}
