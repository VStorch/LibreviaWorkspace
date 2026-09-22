import { describe, expect, it } from 'vitest'
import {
  MIN_IMAGE_PX,
  RESIZE_HANDLES,
  ResizeHandle,
  fittedImage,
  isCornerHandle,
  resizedImage,
} from './image-resize.js'

const start = { width: 400, height: 300 }
const wide = 624

describe('redimensionar imagem pelas alças', () => {
  it('o canto de baixo à direita cresce nas duas medidas', () => {
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 100,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    // A proporção travada faz a altura acompanhar a largura mesmo sem o
    // ponteiro ter subido ou descido: é o eixo horizontal que manda.
    expect(size).toEqual({ width: 500, height: 375 })
  })

  it('o canto de cima à esquerda cresce quando o ponteiro vai para trás', () => {
    const size = resizedImage({
      handle: ResizeHandle.NorthWest,
      start,
      deltaX: -100,
      deltaY: -100,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 375 })
  })

  it('com a proporção solta os dois eixos andam sozinhos', () => {
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 100,
      deltaY: -100,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 200 })
  })

  it('a alça de borda mexe numa medida só, mesmo com a proporção pedida', () => {
    // A alça do meio existe justamente para esticar num eixo: travar a
    // proporção nela a tornaria igual à do canto.
    const size = resizedImage({
      handle: ResizeHandle.East,
      start,
      deltaX: 100,
      deltaY: 250,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 300 })
  })

  it('a alça de baixo mexe só na altura', () => {
    const size = resizedImage({
      handle: ResizeHandle.South,
      start,
      deltaX: 250,
      deltaY: 60,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 400, height: 360 })
  })

  it('a imagem não passa da coluna de texto', () => {
    // Sem o teto, o gravador encolheria a imagem por conta própria ao salvar — e
    // o que a tela mostra deixaria de ser o que o arquivo tem.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 900,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: wide, height: 468 })
  })

  it('a imagem não encolhe até não dar para pegar de volta', () => {
    // Abaixo do mínimo as oito alças se sobrepõem, e crescer de volta só seria
    // possível desfazendo um arrasto que a pessoa não terminou.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: -1000,
      deltaY: -1000,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: MIN_IMAGE_PX, height: MIN_IMAGE_PX })
  })

  it('o tamanho sai em pixel inteiro', () => {
    // O número vai para o atributo do nó, e um valor com casas decimais mudaria
    // a cada quadro por ruído de ponto flutuante — a paginação remede em todos.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start: { width: 401, height: 297 },
      deltaX: 3,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(Number.isInteger(size.width)).toBe(true)
    expect(Number.isInteger(size.height)).toBe(true)
  })

  it('oito alças: quatro cantos e quatro bordas', () => {
    expect(RESIZE_HANDLES).toHaveLength(8)
    expect(RESIZE_HANDLES.filter(isCornerHandle)).toHaveLength(4)
  })
})

describe('imagem que deixou de caber', () => {
  it('a coluna que encolhe encolhe a imagem na proporção', () => {
    expect(fittedImage({ width: 800, height: 600 }, 400)).toEqual({ width: 400, height: 300 })
  })

  it('a imagem que já cabe fica como está', () => {
    const size = { width: 300, height: 200 }
    expect(fittedImage(size, 400)).toBe(size)
  })
})
