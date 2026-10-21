import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { containedPath } from './contained.js'

const ROOT = join('/', 'app', 'out', 'renderer')

describe('caminho contido numa pasta', () => {
  it('accepts a file inside the folder', () => {
    expect(containedPath(ROOT, 'pdf-viewer/index.html')).toBe(join(ROOT, 'pdf-viewer', 'index.html'))
  })

  it('refuses going up a folder, absolute paths and empty names', () => {
    expect(containedPath(ROOT, '../main/index.js')).toBeNull()
    expect(containedPath(ROOT, 'assets/../../main/index.js')).toBeNull()
    expect(containedPath(ROOT, '/etc/passwd')).toBeNull()
    expect(containedPath(ROOT, '')).toBeNull()
  })
})
