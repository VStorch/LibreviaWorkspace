import { readFile } from 'node:fs/promises'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'

/**
 * Lê um arquivo de texto respeitando a marca de ordem de bytes (BOM).
 *
 * Ler tudo como UTF-8 embaralharia arquivos gravados pelo Bloco de Notas do
 * Windows, que são comuns em ambiente corporativo. Tratar o BOM resolve os
 * casos frequentes; detecção de codificação sem BOM (latin-1, por exemplo)
 * fica para quando aparecer um arquivo real que precise disso.
 */
export async function readTextFile(path: string): Promise<string> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  if (startsWith(bytes, [0xef, 0xbb, 0xbf])) {
    return bytes.subarray(3).toString('utf8')
  }

  if (startsWith(bytes, [0xff, 0xfe])) {
    return bytes.subarray(2).toString('utf16le')
  }

  if (startsWith(bytes, [0xfe, 0xff])) {
    return swapByteOrder(bytes.subarray(2)).toString('utf16le')
  }

  assertNotBinary(bytes)
  return bytes.toString('utf8')
}

function startsWith(bytes: Buffer, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false
  return prefix.every((byte, index) => bytes[index] === byte)
}

/**
 * Byte zero é o sinal mais confiável de conteúdo binário em texto sem BOM.
 * Sem esta checagem, abrir um executável por engano encheria o editor de lixo
 * — e salvá-lo destruiria o arquivo original.
 */
function assertNotBinary(bytes: Buffer): void {
  const sample = bytes.subarray(0, 8192)
  if (sample.includes(0)) {
    throw new AppError(
      ErrorCode.NotTextFile,
      'Este arquivo não parece ser de texto e não pode ser aberto com segurança.',
    )
  }
}

function swapByteOrder(bytes: Buffer): Buffer {
  const swapped = Buffer.from(bytes)
  // `swap16` exige comprimento par; um byte solto no fim é lixo e é descartado.
  return swapped.subarray(0, swapped.length - (swapped.length % 2)).swap16()
}
