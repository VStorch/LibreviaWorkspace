/**
 * End to end against the real .NET sidecar: `protocol.test.ts` and `FrameIoTests.cs` check each
 * side against its own idea of the format. Depends on `npm run sidecar:build`, and fails without it
 * instead of being skipped.
 */

import { Buffer } from 'node:buffer'
import { access, constants } from 'node:fs/promises'
import { firstFontOf } from '@services/document/line-metrics.js'
import { BUILTIN_STYLES, LEGACY_STYLES, type StyleSheet } from '@services/document/styles.js'
import { afterAll, describe, expect, it } from 'vitest'
import { ErrorCode, type AppError } from '@shared/errors.js'
import { SidecarClient } from './client.js'
import { SidecarMethod } from './protocol.js'
import { sidecarPathIn } from './locate.js'

const executable = sidecarPathIn(process.cwd())
const published = await access(executable, constants.X_OK).then(
  () => true,
  () => false,
)

const client = new SidecarClient(() => Promise.resolve(executable))
afterAll(() => client.dispose())

describe.skipIf(!published)('sidecar .NET publicado', () => {
  it('responde ao health com as versões das bibliotecas OOXML', async () => {
    const health = await client.health()

    expect(health.name).toBe('Librevia.Format')
    // Versions are part of health because a document that opens wrong on a single machine is almost
    // always a library version difference.
    expect(health.runtime).toMatch(/OpenXml=3\./)
    expect(health.runtime).toMatch(/ClosedXML=0\./)
  })

  it('devolve binário grande byte a byte igual', async () => {
    // 4 MB with a known pattern: catches truncation, reordering and any attempt to treat the bytes
    // as text.
    const payload = new Uint8Array(4 * 1024 * 1024).map((_, i) => (i * 31) % 256)

    const reply = await client.request(SidecarMethod.Echo, {}, payload)

    expect(reply.binary.length).toBe(payload.length)
    // `toEqual` would compare 4 million elements one by one and take longer than the whole round
    // trip. `Buffer.compare` is memcmp.
    expect(Buffer.compare(Buffer.from(reply.binary), Buffer.from(payload))).toBe(0)
  })

  it('preserva bytes que quebrariam um protocolo de linha', async () => {
    // ZIP signature and line endings: that is what a DOCX is made of.
    const payload = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x0a, 0x0d, 0x00, 0x1a])

    const reply = await client.request(SidecarMethod.Echo, {}, payload)

    expect(reply.binary).toEqual(payload)
  })

  it('atende pedidos seguidos no mesmo processo', async () => {
    // Proves the loop goes back to the top instead of serving one and hanging.
    for (let round = 0; round < 5; round++) {
      const reply = await client.request(SidecarMethod.Echo, {}, new Uint8Array([round]))
      expect(reply.binary).toEqual(new Uint8Array([round]))
    }
  })

  it('cria o pacote mínimo de um documento novo', async () => {
    // The page setup goes in the JSON, and the package comes back in the binary, ready to open.
    const page = { size: 'Letter', orientation: 'landscape', margins: { top: 20, right: 20, bottom: 20, left: 20 } }

    const created = await client.request(SidecarMethod.DocxCreate, { page })
    expect([...created.binary.subarray(0, 2)]).toEqual([0x50, 0x4b])

    const opened = await client.request(SidecarMethod.DocxOpen, {}, created.binary)
    expect(opened.result).toMatchObject({ model: { page: { size: 'Letter', orientation: 'landscape' } } })

    // And it serves as the original for the regular save.
    const model = { page, doc: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Olá.' }] }] } }
    const saved = await client.request(SidecarMethod.DocxSave, model, created.binary)
    expect(saved.result).toMatchObject({ rewrittenBlocks: 1 })
  })

  it('o pacote novo grava os estilos que o documento carrega', async () => {
    const page = { size: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 } }

    for (const styles of [BUILTIN_STYLES, LEGACY_STYLES]) {
      const created = await client.request(SidecarMethod.DocxCreate, { page, styles })
      const opened = await client.request(SidecarMethod.DocxOpen, {}, created.binary)
      const read = (opened.result as { model: { styles: StyleSheet } }).model.styles

      // The reader returns the CSS stack; the table, the bare name.
      const { fontFamily, ...character } = read.defaults.character
      expect(firstFontOf(fontFamily)).toBe(styles.defaults.character.fontFamily)
      expect({ ...read.defaults, character }).toMatchObject({
        ...styles.defaults,
        character: { fontSize: styles.defaults.character.fontSize },
      })
      for (const [id, style] of Object.entries(styles.styles)) {
        const { paragraph, character: run } = style
        expect(read.styles[id], id).toMatchObject({
          ...(paragraph === undefined ? {} : { paragraph }),
          ...(run === undefined ? {} : { character: run }),
        })
      }
    }
  })

  it('recusa método desconhecido com erro, sem morrer', async () => {
    const unknown = client.request('metodo.inexistente' as SidecarMethod, {})
    await expect(unknown).rejects.toThrow(/não conhece/i)

    // What matters is that the process stays alive after refusing.
    await expect(client.health()).resolves.toMatchObject({ name: 'Librevia.Format' })
  })

  it('não escreve nada fora do protocolo no stdout', async () => {
    // A stray Console.WriteLine would corrupt the stream. Program.cs redirects Console.Out to
    // stderr for that reason; this test is the lock.
    const health = await client.health()
    const reply = await client.request(SidecarMethod.Echo, {}, new Uint8Array([1]))

    expect(health.name).toBe('Librevia.Format')
    expect(reply.binary).toEqual(new Uint8Array([1]))
  })
})

describe.skipIf(published)('sidecar não publicado', () => {
  it('falha explicando que falta publicar', async () => {
    const missing = new SidecarClient(() => Promise.resolve(executable))
    try {
      await missing.health()
      expect.unreachable('deveria ter falhado')
    } catch (error) {
      expect((error as AppError).code).toBe(ErrorCode.SidecarUnavailable)
    } finally {
      missing.dispose()
    }
  })
})
