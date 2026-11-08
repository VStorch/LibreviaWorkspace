import { randomUUID } from 'node:crypto'
import { mkdtemp, readdir, rm, writeFile, mkdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import { discardDraft, listDraftSummaries, readDraft, useRecoveryFolder, writeDraft } from './recovery.js'

let directory: string
let id: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-recovery-'))
  useRecoveryFolder(directory)
  id = randomUUID()
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

function draftsFolder(): string {
  return join(directory, 'recovery', 'drafts')
}

async function writeRaw(name: string, text: string): Promise<void> {
  await mkdir(draftsFolder(), { recursive: true })
  await writeFile(join(draftsFolder(), name), text)
}

describe('recovery draft', () => {
  it('writes and reads back', async () => {
    const savedAt = await writeDraft(id, SAMPLE)
    const draft = await readDraft(id)

    expect(draft).toEqual({ ...SAMPLE, draftId: id, savedAt })
  })

  it('stores each draft as drafts/<draftId>.json', async () => {
    await writeDraft(id, SAMPLE)

    expect(await readdir(draftsFolder())).toEqual([`${id}.json`])
  })

  it('there is no draft when none was ever written', async () => {
    expect(await readDraft(id)).toBeNull()
    expect(await listDraftSummaries()).toEqual([])
  })

  it('keeps a single draft per id', async () => {
    await writeDraft(id, SAMPLE)
    await writeDraft(id, { ...SAMPLE, name: 'outra.ssheet', kind: DocumentKind.Spreadsheet })

    expect((await readDraft(id))?.name).toBe('outra.ssheet')
    expect(await listDraftSummaries()).toHaveLength(1)
  })

  it('keeps the drafts of different ids apart', async () => {
    const other = randomUUID()
    await writeDraft(id, SAMPLE)
    await writeDraft(other, { ...SAMPLE, name: 'outra.sdoc' })

    expect((await readDraft(id))?.name).toBe('ata.docx')
    expect((await readDraft(other))?.name).toBe('outra.sdoc')
  })

  it('lists the most recent draft first', async () => {
    const older = randomUUID()
    await writeRaw(`${older}.json`, JSON.stringify({ ...SAMPLE, name: 'velho.sdoc', savedAt: 1000 }))
    await writeRaw(`${id}.json`, JSON.stringify({ ...SAMPLE, savedAt: 2000 }))

    expect((await listDraftSummaries()).map((draft) => draft.draftId)).toEqual([id, older])
  })

  it('leaves no .bak copy behind', async () => {
    // The draft is rewritten every eight seconds: keeping the previous version of each would double
    // the writes without protecting anything.
    await writeDraft(id, SAMPLE)
    await writeDraft(id, SAMPLE)

    expect(await readdir(draftsFolder())).toEqual([`${id}.json`])
  })

  it('the summary does not carry the content', async () => {
    // A `.sdoc` with embedded images is tens of megabytes, and the prompt only needs the name and
    // time.
    await writeDraft(id, { ...SAMPLE, content: 'x'.repeat(100_000) })

    expect((await listDraftSummaries())[0]).not.toHaveProperty('content')
  })

  it('a corrupt draft counts as missing', async () => {
    // It exists to save the day after a crash. A read error in it would be a second failure on top
    // of someone who just lost work.
    await writeRaw(`${id}.json`, '{ this is not json')

    expect(await readDraft(id)).toBeNull()
    expect(await listDraftSummaries()).toEqual([])
  })

  it('a draft in an unknown format counts as missing', async () => {
    await writeRaw(`${id}.json`, '{"name":"without the other fields"}')

    expect(await readDraft(id)).toBeNull()
  })

  it('ignores files in the drafts folder that are not named by an id', async () => {
    await writeRaw('notes.json', JSON.stringify({ ...SAMPLE, savedAt: 1000 }))

    expect(await listDraftSummaries()).toEqual([])
  })

  it('discarding deletes', async () => {
    await writeDraft(id, SAMPLE)
    await discardDraft(id)

    expect(await readDraft(id)).toBeNull()
  })

  it('discarding what does not exist is not an error', async () => {
    await expect(discardDraft(id)).resolves.toBeUndefined()
  })

  it('work that was never saved also has a draft', async () => {
    // Exactly the case where recovery matters most: there is no file to go back to.
    await writeDraft(id, { ...SAMPLE, path: null, name: 'Untitled document.sdoc' })

    expect((await readDraft(id))?.path).toBeNull()
  })

  it.each(['../draft', '../../settings', 'not-a-uuid', ''])('rejects the id %j', async (badId) => {
    // The id becomes a file name: anything but a UUID could climb out of the drafts folder.
    await expect(writeDraft(badId, SAMPLE)).rejects.toThrow()
    await expect(readDraft(badId)).rejects.toThrow()
    await expect(discardDraft(badId)).rejects.toThrow()
  })
})

describe('legacy single draft', () => {
  async function writeLegacy(text: string): Promise<void> {
    await mkdir(join(directory, 'recovery'), { recursive: true })
    await writeFile(join(directory, 'recovery', 'draft.json'), text)
  }

  it('moves draft.json into drafts/ under a fresh id on the first read', async () => {
    await writeLegacy(JSON.stringify({ ...SAMPLE, savedAt: 1000 }))

    const [summary] = await listDraftSummaries()

    expect(summary).toMatchObject({ name: 'ata.docx', savedAt: 1000 })
    expect(await readdir(join(directory, 'recovery'))).toEqual(['drafts'])
    expect(await readdir(draftsFolder())).toEqual([`${summary?.draftId}.json`])
    expect((await readDraft(summary?.draftId ?? ''))?.content).toBe(SAMPLE.content)
  })

  it('migrates only once', async () => {
    await writeLegacy(JSON.stringify({ ...SAMPLE, savedAt: 1000 }))

    await listDraftSummaries()

    expect(await listDraftSummaries()).toHaveLength(1)
  })

  it('a corrupt legacy draft still counts as missing after the move', async () => {
    await writeLegacy('{ this is not json')

    expect(await listDraftSummaries()).toEqual([])
  })
})
