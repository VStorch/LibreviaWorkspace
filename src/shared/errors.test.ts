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
    // The difference changes what the user does: in the first case the disk has room, and what ran
    // out is their quota on the network share.
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
    // The system code says nothing to someone who only wants to know whether they can keep working.
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
    // Stack traces and absolute paths do not cross IPC: the renderer is untrusted, and diagnostics
    // stay in the main log.
    const leaked = toSerializedError(new Error('/home/ana/segredo.docx ENOENT at Object.<anonymous>'))

    expect(leaked.code).toBe(ErrorCode.Internal)
    expect(leaked.message).not.toContain('/home/ana')
  })
})
