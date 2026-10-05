import { mkdir, readdir } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, BrowserWindow, shell, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { MessageKey } from '@shared/i18n/index.js'
import type { TemplateEntry } from '@shared/types.js'
import { isWordTemplatePath } from '@services/file/formats.js'
import { showTemplatePickerDialog } from '../dialogs.js'
import { normalizePath } from '../fs/paths.js'
import { t } from '../i18n.js'
import { loadFile } from './file.js'
import { handle } from './registry.js'
import { MAX_USER_TEMPLATES } from '@shared/limits.js'

/**
 * Builtin templates live in `resources/templates`; the user's in `<userData>/Modelos`, where "save
 * as template" puts them. Opening goes through the same `loadFile`.
 */

interface BuiltinTemplate {
  readonly file: string
  readonly name: MessageKey
  readonly description: MessageKey
}

/** The blank document first, as in Word. */
const BUILTIN_TEMPLATES: readonly BuiltinTemplate[] = [
  {
    file: 'documento-em-branco.dotx',
    name: 'shell.template.blank.name',
    description: 'shell.template.blank.description',
  },
  {
    file: 'carta.dotx',
    name: 'shell.template.letter.name',
    description: 'shell.template.letter.description',
  },
  {
    file: 'relatorio.dotx',
    name: 'shell.template.report.name',
    description: 'shell.template.report.description',
  },
  {
    file: 'ata-reuniao.dotx',
    name: 'shell.template.minutes.name',
    description: 'shell.template.minutes.description',
  },
]

/** Like `fonts.ts`: packaged it is `process.resourcesPath`. */
function builtinRoot(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'templates')
}

export function userTemplatesFolder(): string {
  return join(app.getPath('userData'), 'Modelos')
}

function builtinEntries(): TemplateEntry[] {
  return BUILTIN_TEMPLATES.map(({ file, name, description }) => ({
    source: 'builtin',
    id: file,
    name: t(name),
    description: t(description),
  }))
}

/** Without descending into subfolders. */
async function userEntries(): Promise<TemplateEntry[]> {
  const folder = userTemplatesFolder()
  let names: string[]
  try {
    const entries = await readdir(folder, { withFileTypes: true })
    names = entries
      .filter((entry) => entry.isFile() && isWordTemplatePath(entry.name))
      .map((entry) => entry.name)
  } catch {
    // Nobody has saved a template yet.
    return []
  }

  return names
    .sort((a, b) => a.localeCompare(b, 'pt-BR'))
    .slice(0, MAX_USER_TEMPLATES)
    .map((name) => ({
      source: 'user',
      id: join(folder, name),
      name: name.replace(/\.dot[xm]$/i, ''),
      description: t('shell.template.userTemplate'),
    }))
}

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, t('errors.ipc.windowNotAvailable'))
  }
  return window
}

export function registerTemplateHandlers(): void {
  handle(IpcChannel.TemplateList, async () => ({
    builtin: builtinEntries(),
    user: await userEntries(),
    folder: userTemplatesFolder(),
  }))

  handle(IpcChannel.TemplateOpen, async (payload) => {
    // The id must be in the list main builds itself.
    if (payload.source === 'builtin') {
      const builtin = builtinEntries().find((entry) => entry.id === payload.id)
      if (builtin === undefined)
        throw new AppError(ErrorCode.PathNotAuthorized, t('errors.paths.unauthorized'))
      return { file: await loadFile(join(builtinRoot(), builtin.id), builtin.name) }
    }

    const wanted = normalizePath(payload.id)
    const user = (await userEntries()).find((entry) => normalizePath(entry.id) === wanted)
    if (user === undefined) throw new AppError(ErrorCode.PathNotAuthorized, t('errors.paths.unauthorized'))
    return { file: await loadFile(user.id) }
  })

  handle(IpcChannel.TemplateBrowse, async (_payload, event) => {
    const path = await showTemplatePickerDialog(windowOf(event), userTemplatesFolder())
    if (path === null) return { canceled: true as const }
    if (!isWordTemplatePath(path)) {
      throw new AppError(ErrorCode.UnsupportedFormat, t('errors.paths.unsupportedType'))
    }
    return { canceled: false as const, file: await loadFile(path) }
  })

  handle(IpcChannel.TemplateOpenFolder, async () => {
    const folder = userTemplatesFolder()
    await mkdir(folder, { recursive: true })
    // `openPath` returns the error message instead of throwing.
    const problem = await shell.openPath(folder)
    if (problem !== '') console.warn(`[templates] ${problem}`)
    return { folder }
  })
}
