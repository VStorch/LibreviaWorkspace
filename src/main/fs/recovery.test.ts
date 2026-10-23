import { mkdtemp, readdir, rm, writeFile, mkdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import { discardDraft, readDraft, readDraftSummary, useRecoveryFolder, writeDraft } from './recovery.js'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-recovery-'))
  useRecoveryFolder(directory)
})

afterEach(async () => {
  await rm(directory, { recursive: true, force: true })
})

const SAMPLE = {
  path: '/home/ana/ata.docx',
  name: 'ata.docx',
  kind: DocumentKind.Document,
  content: '{"format":"sdoc"}',
}

describe('recovery draft', () => {
  it('writes and reads back', async () => {
    const savedAt = await writeDraft(SAMPLE)
    const draft = await readDraft()

    expect(draft).toEqual({ ...SAMPLE, savedAt })
  })

  it('there is no draft when none was ever written', async () => {
    expect(await readDraft()).toBeNull()
  })

  it('keeps a single draft', async () => {
    await writeDraft(SAMPLE)
    await writeDraft({ ...SAMPLE, name: 'outra.ssheet', kind: DocumentKind.Spreadsheet })

    expect((await readDraft())?.name).toBe('outra.ssheet')
  })

  it('leaves no .bak copy behind', async () => {
    // The draft is rewritten every eight seconds: keeping the previous version of each would double
    // the writes without protecting anything.
    await writeDraft(SAMPLE)
    await writeDraft(SAMPLE)

    expect(await readdir(join(directory, 'recovery'))).toEqual(['draft.json'])
  })

  it('the summary does not carry the content', async () => {
    // A `.sdoc` with embedded images is tens of megabytes, and the prompt only needs the name and
    // time.
    await writeDraft({ ...SAMPLE, content: 'x'.repeat(100_000) })

    expect(await readDraftSummary()).not.toHaveProperty('content')
  })

  it('a corrupt draft counts as missing', async () => {
    // It exists to save the day after a crash. A read error in it would be a second failure on top
    // of someone who just lost work.
    await mkdir(join(directory, 'recovery'), { recursive: true })
    await writeFile(join(directory, 'recovery', 'draft.json'), '{ this is not json')

    expect(await readDraft()).toBeNull()
  })

  it('a draft in an unknown format counts as missing', async () => {
    await mkdir(join(directory, 'recovery'), { recursive: true })
    await writeFile(join(directory, 'recovery', 'draft.json'), '{"name":"without the other fields"}')

    expect(await readDraft()).toBeNull()
  })

  it('discarding deletes', async () => {
    await writeDraft(SAMPLE)
    await discardDraft()

    expect(await readDraft()).toBeNull()
  })

  it('discarding what does not exist is not an error', async () => {
    await expect(discardDraft()).resolves.toBeUndefined()
  })

  it('work that was never saved also has a draft', async () => {
    // Exactly the case where recovery matters most: there is no file to go back to.
    await writeDraft({ ...SAMPLE, path: null, name: 'Untitled document.sdoc' })

    expect((await readDraft())?.path).toBeNull()
  })
})
