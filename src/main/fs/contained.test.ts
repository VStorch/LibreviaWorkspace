import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { containedPath } from './contained.js'

const ROOT = join('/', 'app', 'out', 'renderer')

describe('caminho contido numa pasta', () => {
  it('aceita um arquivo dentro da pasta', () => {
    expect(containedPath(ROOT, 'pdf-viewer/index.html')).toBe(join(ROOT, 'pdf-viewer', 'index.html'))
  })

  it('recusa subir de pasta, caminho absoluto e nome vazio', () => {
    expect(containedPath(ROOT, '../main/index.js')).toBeNull()
    expect(containedPath(ROOT, 'assets/../../main/index.js')).toBeNull()
    expect(containedPath(ROOT, '/etc/passwd')).toBeNull()
    expect(containedPath(ROOT, '')).toBeNull()
  })
})
