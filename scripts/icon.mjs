/**
 * Gera o ícone do aplicativo — um PNG escrito à mão, sem dependência nenhuma.
 *
 * É um **provisório assumido**: quadrado de canto arredondado com um "L" claro.
 * Existe porque o instalador precisa de um ícone e porque um ícone de verdade é
 * trabalho de quem desenha, não de quem programa. Trocar é só substituir
 * `build/icon.png` por um PNG de 512×512.
 *
 * Escrever o PNG à mão em vez de instalar uma biblioteca de imagem: o formato
 * mínimo são três blocos e um CRC, e `node:zlib` já faz a única parte difícil.
 */

import { deflateSync } from 'node:zlib'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'

const SIZE = 512
const RADIUS = 96
const BACKGROUND = [26, 58, 92] // azul-ardósia
const MARK = [244, 246, 249]

/** O "L", em fração do lado: haste vertical e pé horizontal. */
const STEM = { x: 0.34, y: 0.24, w: 0.1, h: 0.52 }
const FOOT = { x: 0.34, y: 0.66, w: 0.32, h: 0.1 }

function inside(x, y, box) {
  return x >= box.x * SIZE && x < (box.x + box.w) * SIZE && y >= box.y * SIZE && y < (box.y + box.h) * SIZE
}

/** Canto arredondado: fora do raio, o pixel é transparente. */
function opaque(x, y) {
  const corners = [
    [RADIUS, RADIUS],
    [SIZE - RADIUS - 1, RADIUS],
    [RADIUS, SIZE - RADIUS - 1],
    [SIZE - RADIUS - 1, SIZE - RADIUS - 1],
  ]

  const nearLeft = x < RADIUS
  const nearRight = x > SIZE - RADIUS - 1
  const nearTop = y < RADIUS
  const nearBottom = y > SIZE - RADIUS - 1
  if (!((nearLeft || nearRight) && (nearTop || nearBottom))) return true

  const [cx, cy] = corners.find(
    ([ax, ay]) => x < RADIUS === (ax === RADIUS) && y < RADIUS === (ay === RADIUS),
  )
  return (x - cx) ** 2 + (y - cy) ** 2 <= RADIUS ** 2
}

function pixels() {
  // Uma linha de filtro (0 = nenhum) antes de cada linha de pixels, como o PNG pede.
  const raw = Buffer.alloc(SIZE * (SIZE * 4 + 1))
  let at = 0

  for (let y = 0; y < SIZE; y++) {
    raw[at++] = 0
    for (let x = 0; x < SIZE; x++) {
      const visible = opaque(x, y)
      const mark = inside(x, y, STEM) || inside(x, y, FOOT)
      const [r, g, b] = mark ? MARK : BACKGROUND
      raw[at++] = r
      raw[at++] = g
      raw[at++] = b
      raw[at++] = visible ? 255 : 0
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

function png() {
  const header = Buffer.alloc(13)
  header.writeUInt32BE(SIZE, 0)
  header.writeUInt32BE(SIZE, 4)
  header[8] = 8 // bits por canal
  header[9] = 6 // RGBA
  // profundidade, filtro e entrelaçamento ficam em zero: o básico do formato.

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(pixels(), { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

const target = resolve('build/icon.png')
mkdirSync(dirname(target), { recursive: true })
writeFileSync(target, png())
console.log(`ícone gravado em ${target} (${SIZE}×${SIZE})`)
