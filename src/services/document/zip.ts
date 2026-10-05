/**
 * ODF requires `mimetype` as the **first** entry, uncompressed and without an extra field.
 * Compression comes from outside (`zlib` in main); without it everything is stored, still a valid
 * ZIP.
 */

export interface ZipEntry {
  readonly name: string
  readonly data: Uint8Array
  /** Stored as is: `mimetype` and images, which come already compressed. */
  readonly stored?: boolean
}

export type Deflate = (data: Uint8Array) => Uint8Array

const STORED = 0
const DEFLATED = 8
/** Bit 11: UTF-8 name. */
const UTF8_FLAG = 0x0800

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of data) crc = CRC_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

/** In MS-DOS format, which ZIP uses; nothing before 1980 exists. */
function dosDateTime(date: Date): { time: number; date: number } {
  const year = Math.min(Math.max(date.getFullYear(), 1980), 2107)
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  }
}

/**
 * Without `deflate` everything is stored; the date is fixed by default, for the same bytes in
 * tests.
 */
export function zip(
  entries: readonly ZipEntry[],
  deflate?: Deflate,
  modified: Date = new Date(1980, 0, 1),
): Uint8Array {
  const encoder = new TextEncoder()
  const stamp = dosDateTime(modified)
  const locals: Uint8Array[] = []
  const centrals: Uint8Array[] = []
  let offset = 0

  for (const entry of entries) {
    const name = encoder.encode(entry.name)
    const crc = crc32(entry.data)
    let method = STORED
    let body = entry.data
    if (entry.stored !== true && deflate !== undefined) {
      const packed = deflate(entry.data)
      if (packed.length < entry.data.length) {
        method = DEFLATED
        body = packed
      }
    }

    const local = new Uint8Array(30 + name.length)
    const head = new DataView(local.buffer)
    head.setUint32(0, 0x04034b50, true)
    head.setUint16(4, method === DEFLATED ? 20 : 10, true)
    head.setUint16(6, UTF8_FLAG, true)
    head.setUint16(8, method, true)
    head.setUint16(10, stamp.time, true)
    head.setUint16(12, stamp.date, true)
    head.setUint32(14, crc, true)
    head.setUint32(18, body.length, true)
    head.setUint32(22, entry.data.length, true)
    head.setUint16(26, name.length, true)
    head.setUint16(28, 0, true)
    local.set(name, 30)

    const central = new Uint8Array(46 + name.length)
    const dir = new DataView(central.buffer)
    dir.setUint32(0, 0x02014b50, true)
    dir.setUint16(4, 20, true)
    dir.setUint16(6, method === DEFLATED ? 20 : 10, true)
    dir.setUint16(8, UTF8_FLAG, true)
    dir.setUint16(10, method, true)
    dir.setUint16(12, stamp.time, true)
    dir.setUint16(14, stamp.date, true)
    dir.setUint32(16, crc, true)
    dir.setUint32(20, body.length, true)
    dir.setUint32(24, entry.data.length, true)
    dir.setUint16(28, name.length, true)
    dir.setUint32(42, offset, true)
    central.set(name, 46)

    locals.push(local, body)
    centrals.push(central)
    offset += local.length + body.length
  }

  const directorySize = centrals.reduce((sum, part) => sum + part.length, 0)
  const end = new Uint8Array(22)
  const tail = new DataView(end.buffer)
  tail.setUint32(0, 0x06054b50, true)
  tail.setUint16(8, entries.length, true)
  tail.setUint16(10, entries.length, true)
  tail.setUint32(12, directorySize, true)
  tail.setUint32(16, offset, true)

  const parts = [...locals, ...centrals, end]
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}
