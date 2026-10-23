import { describe, expect, it } from 'vitest'
import { ErrorCode, fromFileSystemError, toSerializedError, AppError } from './errors.js'

describe('fromFileSystemError', () => {
  const of = (code: string, operation: 'read' | 'write' = 'write') =>
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
  ])('%s becomes %s', (errno, expected) => {
    expect(of(errno).code).toBe(expected)
  })

  it('tells an exhausted quota from a full disk', () => {
    // The difference changes what the user does: in the first case the disk has room, and what ran
    // out is their quota on the network share.
    expect(of('ENOSPC').message).not.toBe(of('EDQUOT').message)
    expect(of('EDQUOT').message).toContain('cota')
  })

  it.each(['ENETDOWN', 'EHOSTUNREACH', 'ESTALE', 'ETIMEDOUT'])(
    'treats %s as a network folder that went down',
    (errno) => {
      expect(of(errno).message).toContain('rede')
    },
  )

  it('the operation decides the code of an unknown error', () => {
    expect(of('EWHATEVER', 'read').code).toBe(ErrorCode.ReadFailed)
    expect(of('EWHATEVER', 'write').code).toBe(ErrorCode.WriteFailed)
  })

  it('no message carries the raw errno', () => {
    // The system code says nothing to someone who only wants to know whether they can keep working.
    for (const errno of ['ENOENT', 'EACCES', 'EDQUOT', 'EBUSY', 'EXDEV']) {
      expect(of(errno).message).not.toContain(errno)
    }
  })
})

describe('toSerializedError', () => {
  it('keeps the message of an AppError', () => {
    const error = new AppError(ErrorCode.FileTooLarge, 'O arquivo é grande demais.', 'limite 20 MB')

    expect(toSerializedError(error)).toEqual({
      code: ErrorCode.FileTooLarge,
      message: 'O arquivo é grande demais.',
      detail: 'limite 20 MB',
    })
  })

  it('does not leak what was thrown by accident', () => {
    // Stack traces and absolute paths do not cross IPC: the renderer is untrusted, and diagnostics
    // stay in the main log.
    const leaked = toSerializedError(new Error('/home/ana/segredo.docx ENOENT at Object.<anonymous>'))

    expect(leaked.code).toBe(ErrorCode.Internal)
    expect(leaked.message).not.toContain('/home/ana')
  })
})
