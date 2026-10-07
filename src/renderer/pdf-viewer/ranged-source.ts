import { PDFDataRangeTransport } from 'pdfjs-dist'

/** pdf.js's own chunk size: the header, and all of a small PDF, arrive in the first one. */
const CHUNK_BYTES = 65_536
const RANGE_NOT_SATISFIABLE = 416

export type PdfSource =
  { readonly data: Uint8Array } | { readonly range: PDFDataRangeTransport; readonly length: number }

/**
 * pdf.js asks for ranges only over http(s); from any other scheme it reads the whole file. Through
 * a transport it asks for the pieces it needs, and a PDF of hundreds of megabytes opens without
 * reaching memory whole. `onFailure` runs if a later piece cannot be read.
 */
export async function rangedSource(url: string, onFailure: () => void): Promise<PdfSource> {
  const first = await fetchRange(url, 0, CHUNK_BYTES)
  // An empty file has no range to give; pdf.js then explains that it is not a PDF.
  if (!first.ok && first.status !== RANGE_NOT_SATISFIABLE) throw new Error(`HTTP ${first.status}`)

  const initial = new Uint8Array(await first.arrayBuffer())
  const length = totalOf(first.headers.get('Content-Range'))
  if (length === null || length <= initial.length) return { data: initial }

  return { range: new FetchRangeTransport(url, length, initial, onFailure), length }
}

class FetchRangeTransport extends PDFDataRangeTransport {
  readonly #url: string
  readonly #onFailure: () => void

  constructor(url: string, length: number, initial: Uint8Array, onFailure: () => void) {
    super(length, initial)
    this.#url = url
    this.#onFailure = onFailure
  }

  override requestDataRange(begin: number, end: number): void {
    fetchRange(this.#url, begin, end)
      .then((response) => response.arrayBuffer())
      .then((chunk) => this.onDataRange(begin, new Uint8Array(chunk)), this.#onFailure)
  }
}

/** `end` is exclusive, as pdf.js counts; HTTP's is inclusive. */
function fetchRange(url: string, begin: number, end: number): Promise<Response> {
  return fetch(url, { headers: { Range: `bytes=${begin}-${end - 1}` } })
}

function totalOf(contentRange: string | null): number | null {
  const total = Number(contentRange?.split('/')[1])
  return Number.isSafeInteger(total) ? total : null
}
