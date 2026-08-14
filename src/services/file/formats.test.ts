import { describe, expect, it } from 'vitest'
import { buildWindowTitle, extensionOf, fileNameFromPath, isSupportedExtension } from './formats.js'

describe('fileNameFromPath', () => {
  it.each([
    ['/home/ana/relatorio.txt', 'relatorio.txt'],
    ['C:\\Users\\Ana\\relatorio.txt', 'relatorio.txt'],
    // Pastas de rede são o caso que a Fase 1 precisa acertar.
    ['\\\\servidor\\setor\\ata.txt', 'ata.txt'],
    ['/mnt/rede/contratos/minuta final.txt', 'minuta final.txt'],
    ['sem-pasta.txt', 'sem-pasta.txt'],
  ])('extrai o nome de %s', (path, expected) => {
    expect(fileNameFromPath(path)).toBe(expected)
  })
})

describe('extensionOf', () => {
  it.each([
    ['/a/b/c.txt', '.txt'],
    ['/a/b/c.TXT', '.txt'],
    ['/a/b/arquivo.com.ponto.txt', '.txt'],
    ['/a/b/sem-extensao', ''],
    // Um ponto inicial é arquivo oculto, não extensão.
    ['/a/b/.oculto', ''],
  ])('lê a extensão de %s', (path, expected) => {
    expect(extensionOf(path)).toBe(expected)
  })

  it('não confunde ponto de pasta com extensão do arquivo', () => {
    expect(extensionOf('/home/ana/pasta.com.ponto/arquivo')).toBe('')
  })
})

describe('isSupportedExtension', () => {
  it.each(['/a/b.txt', '/a/b.sdoc', '/a/b.docx', '/a/b.DOCX'])('aceita %s', (path) => {
    expect(isSupportedExtension(path)).toBe(true)
  })

  it.each(['/a/b.exe', '/a/b.sh', '/a/b', '/a/b.xlsx'])('recusa %s', (path) => {
    // `.xlsx` só na Fase 7 — abrir uma planilha no editor de texto não
    // produziria nada útil.
    expect(isSupportedExtension(path)).toBe(false)
  })
})

describe('buildWindowTitle', () => {
  it('marca alterações não salvas', () => {
    expect(buildWindowTitle('ata.txt', true, 'Librevia')).toBe('• ata.txt — Librevia')
  })

  it('não marca quando está salvo', () => {
    expect(buildWindowTitle('ata.txt', false, 'Librevia')).toBe('ata.txt — Librevia')
  })

  it('usa rótulo próprio para arquivo ainda sem nome', () => {
    expect(buildWindowTitle(null, false, 'Librevia')).toBe('Sem título — Librevia')
  })
})
