/** Programs some tests need beyond the app; a test without them is skipped, not failed. */

import { execFile } from 'node:child_process'
import { promisify } from 'node:util'

const run = promisify(execFile)

async function isInstalled(program: string, versionFlag: string): Promise<boolean> {
  try {
    await run(program, [versionFlag])
    return true
  } catch {
    return false
  }
}

export const hasPdftotext = (): Promise<boolean> => isInstalled('pdftotext', '-v')
export const hasPdfinfo = (): Promise<boolean> => isInstalled('pdfinfo', '-v')
export const hasSoffice = (): Promise<boolean> => isInstalled('soffice', '--version')

/** Empty while the PDF is still being written, so that `expect.poll` retries. */
export async function pdfText(
  path: string,
  { layout = true, page }: { layout?: boolean; page?: number } = {},
): Promise<string> {
  const onePage = page === undefined ? [] : ['-f', String(page), '-l', String(page)]
  const options = [...onePage, ...(layout ? ['-layout'] : [])]
  try {
    const { stdout } = await run('pdftotext', ['-enc', 'UTF-8', ...options, path, '-'])
    return stdout
  } catch {
    return ''
  }
}

/** The metadata poppler reads, one `Key: value` per line; empty on failure. */
export async function pdfInfo(path: string, { lastPage }: { lastPage?: number } = {}): Promise<string> {
  const pageSizes = lastPage === undefined ? [] : ['-f', '1', '-l', String(lastPage)]
  try {
    const { stdout } = await run('pdfinfo', ['-enc', 'UTF-8', ...pageSizes, path])
    return stdout
  } catch {
    return ''
  }
}
