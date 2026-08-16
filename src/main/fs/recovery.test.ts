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
    // O rascunho é reescrito de oito em oito segundos: guardar a versão
    // anterior de cada uma delas dobraria a escrita sem proteger nada.
    await writeDraft(SAMPLE)
    await writeDraft(SAMPLE)

    expect(await readdir(join(directory, 'recuperacao'))).toEqual(['rascunho.json'])
  })

  it('o resumo não carrega o conteúdo', async () => {
    // Um `.sdoc` com imagens embutidas tem dezenas de megabytes, e o aviso só
    // precisa do nome e da hora para o usuário decidir.
    await writeDraft({ ...SAMPLE, content: 'x'.repeat(100_000) })

    expect(await readDraftSummary()).not.toHaveProperty('content')
  })

  it('rascunho corrompido conta como ausente', async () => {
    // Ele existe para salvar o dia depois de uma queda. Um erro de leitura dele
    // viraria uma segunda falha em cima de quem acabou de perder trabalho.
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
    // É justamente o caso em que a recuperação vale mais: não há arquivo
    // nenhum a que voltar.
    await writeDraft({ ...SAMPLE, path: null, name: 'Documento sem título.sdoc' })

    expect((await readDraft())?.path).toBeNull()
  })
})
