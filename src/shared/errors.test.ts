import { describe, expect, it } from 'vitest'
import { ErrorCode, fromFileSystemError, toSerializedError, AppError } from './errors.js'

describe('fromFileSystemError', () => {
  const of = (code: string, operation: 'leitura' | 'escrita' = 'escrita') =>
    fromFileSystemError(Object.assign(new Error('cru'), { code }), operation)

  it.each([
    ['ENOENT', ErrorCode.FileNotFound],
    ['EACCES', ErrorCode.PermissionDenied],
    ['EPERM', ErrorCode.PermissionDenied],
    ['EISDIR', ErrorCode.NotAFile],
    ['EROFS', ErrorCode.WriteFailed],
    ['ENOSPC', ErrorCode.WriteFailed],
    ['EDQUOT', ErrorCode.WriteFailed],
    ['ENAMETOOLONG', ErrorCode.WriteFailed],
    ['EBUSY', ErrorCode.WriteFailed],
  ])('%s vira %s', (errno, expected) => {
    expect(of(errno).code).toBe(expected)
  })

  it('separa cota esgotada de disco cheio', () => {
    // A diferença muda o que a pessoa faz: no primeiro caso o disco tem espaço,
    // e quem acabou foi a cota dela na pasta de rede.
    expect(of('ENOSPC').message).not.toBe(of('EDQUOT').message)
    expect(of('EDQUOT').message).toContain('cota')
  })

  it.each(['ENETDOWN', 'EHOSTUNREACH', 'ESTALE', 'ETIMEDOUT'])(
    'trata %s como pasta de rede que caiu',
    (errno) => {
      expect(of(errno).message).toContain('rede')
    },
  )

  it('a operação decide o código do erro desconhecido', () => {
    expect(of('EWHATEVER', 'leitura').code).toBe(ErrorCode.ReadFailed)
    expect(of('EWHATEVER', 'escrita').code).toBe(ErrorCode.WriteFailed)
  })

  it('nenhuma mensagem carrega o errno cru', () => {
    // O código do sistema não diz nada a quem só quer saber se pode continuar
    // trabalhando — e é isso que a tradução existe para evitar.
    for (const errno of ['ENOENT', 'EACCES', 'EDQUOT', 'EBUSY', 'EXDEV']) {
      expect(of(errno).message).not.toContain(errno)
    }
  })
})

describe('toSerializedError', () => {
  it('preserva a frase de um AppError', () => {
    const error = new AppError(ErrorCode.FileTooLarge, 'O arquivo é grande demais.', 'limite 20 MB')

    expect(toSerializedError(error)).toEqual({
      code: ErrorCode.FileTooLarge,
      message: 'O arquivo é grande demais.',
      detail: 'limite 20 MB',
    })
  })

  it('não deixa vazar o que foi lançado por acidente', () => {
    // Stack trace e caminho absoluto não atravessam o IPC: o renderer é tratado
    // como não confiável, e o diagnóstico fica no log do main.
    const leaked = toSerializedError(new Error('/home/ana/segredo.docx ENOENT at Object.<anonymous>'))

    expect(leaked.code).toBe(ErrorCode.Internal)
    expect(leaked.message).not.toContain('/home/ana')
  })
})
