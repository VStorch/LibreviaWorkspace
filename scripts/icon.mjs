/**
 * Renders `build/icon.svg` into `build/icon.png` (512×512, transparent background). The SVG holds
 * only filled polygons, so a small rasterizer is enough and the build needs no image library.
 */

import { deflateSync } from 'node:zlib'
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'

const SIZE = 512
const SAMPLES_PER_AXIS = 4
const OPAQUE = 255

function readPolygons(svg) {
  return [...svg.matchAll(/<path d="M([^"]+?)\s*Z" fill="#([0-9a-f]{6})"\/>/gi)].map(([, path, hex]) => ({
    points: path.split('L').map((pair) => pair.trim().split(/\s+/).map(Number)),
    color: [0, 2, 4].map((at) => parseInt(hex.slice(at, at + 2), 16)),
  }))
}

function contains(points, x, y) {
  let inside = false
  for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
    const [xi, yi] = points[i]
    const [xj, yj] = points[j]
    if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside
  }
  return inside
}

function coverage(points, px, py) {
  let hits = 0
  for (let sy = 0; sy < SAMPLES_PER_AXIS; sy++) {
    for (let sx = 0; sx < SAMPLES_PER_AXIS; sx++) {
      const x = px + (sx + 0.5) / SAMPLES_PER_AXIS
      const y = py + (sy + 0.5) / SAMPLES_PER_AXIS
      if (contains(points, x, y)) hits++
    }
  }
  return hits / SAMPLES_PER_AXIS ** 2
}

function blendOver(pixel, color, alpha) {
  const outAlpha = alpha + pixel[3] * (1 - alpha)
  if (outAlpha === 0) return pixel
  const channel = (i) => (color[i] * alpha + pixel[i] * pixel[3] * (1 - alpha)) / outAlpha
  return [channel(0), channel(1), channel(2), outAlpha]
}

function pixels(polygons) {
  // PNG wants a filter byte (0 = none) before each row.
  const raw = Buffer.alloc(SIZE * (SIZE * 4 + 1))
  let at = 0
  for (let y = 0; y < SIZE; y++) {
    raw[at++] = 0
    for (let x = 0; x < SIZE; x++) {
      const pixel = polygons.reduce(
        (under, { points, color }) => {
          const alpha = coverage(points, x, y)
          return alpha === 0 ? under : blendOver(under, color, alpha)
        },
        [0, 0, 0, 0],
      )
      raw[at++] = Math.round(pixel[0])
      raw[at++] = Math.round(pixel[1])
      raw[at++] = Math.round(pixel[2])
      raw[at++] = Math.round(pixel[3] * OPAQUE)
    }
  }
  return raw
}

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})

function crc32(buffer) {
  let c = 0xffffffff
  for (const byte of buffer) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

function chunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}

function png(raw) {
  const header = Buffer.alloc(13)
  header.writeUInt32BE(SIZE, 0)
  header.writeUInt32BE(SIZE, 4)
  header[8] = 8 // bits per channel
  header[9] = 6 // RGBA

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

const polygons = readPolygons(readFileSync(resolve('build/icon.svg'), 'utf8'))
if (polygons.length === 0) throw new Error('build/icon.svg has no polygon to draw')
const target = resolve('build/icon.png')
writeFileSync(target, png(pixels(polygons)))
console.log(`icon written to ${target} (${SIZE}×${SIZE}, ${polygons.length} polygons)`)
