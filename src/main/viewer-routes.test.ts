import { describe, expect, it } from 'vitest'
import { contentTypeOf, parseByteRange, viewerRoute } from './viewer-routes.js'

const TOKEN = '0f8b7c2e-41d3-4a5b-9c6d-7e8f9a0b1c2d'

describe('PDF viewer routes', () => {
  it('the PDF is only served by token', () => {
    expect(viewerRoute(`librevia-viewer://app/document/${TOKEN}`)).toEqual({ kind: 'document', token: TOKEN })
    expect(viewerRoute('librevia-viewer://app/document/%2Fhome%2Fana%2Fsegredo.pdf')).toBeNull()
    expect(viewerRoute('librevia-viewer://app/document/')).toBeNull()
  })

  it('serves only the pdf.js resource folders', () => {
    expect(viewerRoute('librevia-viewer://app/pdfjs/cmaps/UniJIS-UCS2-H.bcmap')).toEqual({
      kind: 'pdfjs',
      name: 'cmaps/UniJIS-UCS2-H.bcmap',
    })
    expect(viewerRoute('librevia-viewer://app/pdfjs/build/pdf.mjs')).toBeNull()
    expect(viewerRoute('librevia-viewer://app/pdfjs/package.json')).toBeNull()
  })

  it('serves the viewer page and the bundle, but not the app page', () => {
    expect(viewerRoute('librevia-viewer://app/pdf-viewer/index.html')).toEqual({
      kind: 'page',
      name: 'pdf-viewer/index.html',
    })
    expect(viewerRoute('librevia-viewer://app/assets/viewer-1a2b.js')).toEqual({
      kind: 'page',
      name: 'assets/viewer-1a2b.js',
    })
    expect(viewerRoute('librevia-viewer://app/index.html')).toBeNull()
  })

  it('recusa outro host e outro esquema', () => {
    expect(viewerRoute('librevia-viewer://outro/pdf-viewer/index.html')).toBeNull()
    expect(viewerRoute('librevia-font://app/pdf-viewer/index.html')).toBeNull()
    expect(viewerRoute('not a url')).toBeNull()
  })

  it('an encoded slash passes the route: containment is checked later, on the resolved path', () => {
    expect(viewerRoute('librevia-viewer://app/pdfjs/cmaps/..%2F..%2Fmain%2Findex.js')).toEqual({
      kind: 'pdfjs',
      name: 'cmaps/../../main/index.js',
    })
  })
})

describe('content type', () => {
  it('wasm is served as wasm, so the browser compiles it in streaming', () => {
    expect(contentTypeOf('wasm/openjpeg.wasm')).toBe('application/wasm')
    expect(contentTypeOf('assets/pdf.worker.min-x.mjs')).toBe('text/javascript; charset=utf-8')
    expect(contentTypeOf('cmaps/Adobe-Japan1-0.bcmap')).toBe('application/octet-stream')
  })
})

describe('pedidos por trecho', () => {
  it('without a header it is the whole file', () => {
    expect(parseByteRange(null, 1000)).toBeNull()
  })

  it('reads start and end, and an open end goes to the last byte', () => {
    expect(parseByteRange('bytes=0-99', 1000)).toEqual({ start: 0, end: 99 })
    expect(parseByteRange('bytes=900-', 1000)).toEqual({ start: 900, end: 999 })
    expect(parseByteRange('bytes=900-5000', 1000)).toEqual({ start: 900, end: 999 })
  })

  it('a range outside the file is refused', () => {
    expect(parseByteRange('bytes=1000-1100', 1000)).toBe('unsatisfiable')
    expect(parseByteRange('bytes=50-10', 1000)).toBe('unsatisfiable')
  })

  it('what it does not understand becomes the whole file, as in HTTP', () => {
    expect(parseByteRange('bytes=0-9,20-29', 1000)).toBeNull()
    expect(parseByteRange('items=0-9', 1000)).toBeNull()
  })
})
