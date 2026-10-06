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
export async function pdfText(path: string, { layout = true } = {}): Promise<string> {
  try {
    const { stdout } = await run('pdftotext', [...(layout ? ['-layout'] : []), path, '-'])
    return stdout
  } catch {
    return ''
  }
}
