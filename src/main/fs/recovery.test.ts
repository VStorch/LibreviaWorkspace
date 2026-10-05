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

describe('rascunho de recuperação', () => {
  it('grava e lê de volta', async () => {
    const savedAt = await writeDraft(SAMPLE)
    const draft = await readDraft()

    expect(draft).toEqual({ ...SAMPLE, savedAt })
  })

  it('não há rascunho quando nunca se gravou nenhum', async () => {
    expect(await readDraft()).toBeNull()
  })

  it('guarda um rascunho só', async () => {
    await writeDraft(SAMPLE)
    await writeDraft({ ...SAMPLE, name: 'outra.ssheet', kind: DocumentKind.Spreadsheet })

    expect((await readDraft())?.name).toBe('outra.ssheet')
  })

  it('não deixa cópia .bak para trás', async () => {
    // The draft is rewritten every eight seconds: keeping the previous version of each would double
    // the writes without protecting anything.
    await writeDraft(SAMPLE)
    await writeDraft(SAMPLE)

    expect(await readdir(join(directory, 'recuperacao'))).toEqual(['rascunho.json'])
  })

  it('o resumo não carrega o conteúdo', async () => {
    // A `.sdoc` with embedded images is tens of megabytes, and the prompt only needs the name and
    // time.
    await writeDraft({ ...SAMPLE, content: 'x'.repeat(100_000) })

    expect(await readDraftSummary()).not.toHaveProperty('content')
  })

  it('rascunho corrompido conta como ausente', async () => {
    // It exists to save the day after a crash. A read error in it would be a second failure on top
    // of someone who just lost work.
    await mkdir(join(directory, 'recuperacao'), { recursive: true })
    await writeFile(join(directory, 'recuperacao', 'rascunho.json'), '{ isto não é json')

    expect(await readDraft()).toBeNull()
  })

  it('rascunho de formato desconhecido conta como ausente', async () => {
    await mkdir(join(directory, 'recuperacao'), { recursive: true })
    await writeFile(join(directory, 'recuperacao', 'rascunho.json'), '{"name":"sem os outros campos"}')

    expect(await readDraft()).toBeNull()
  })

  it('descartar apaga', async () => {
    await writeDraft(SAMPLE)
    await discardDraft()

    expect(await readDraft()).toBeNull()
  })

  it('descartar o que não existe não é erro', async () => {
    await expect(discardDraft()).resolves.toBeUndefined()
  })

  it('trabalho que nunca foi gravado também tem rascunho', async () => {
    // Exactly the case where recovery matters most: there is no file to go back to.
    await writeDraft({ ...SAMPLE, path: null, name: 'Documento sem título.sdoc' })

    expect((await readDraft())?.path).toBeNull()
  })
})
