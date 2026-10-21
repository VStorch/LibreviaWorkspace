import { describe, expect, it } from 'vitest'
import { isLoopback } from './locality.js'

describe('isLoopback', () => {
  it.each([
    'http://localhost:11434/v1',
    'http://127.0.0.1:1234/v1',
    'http://127.8.9.10/v1',
    'http://[::1]:8080/v1',
    'http://LOCALHOST/v1',
  ])('treats %s as this computer', (url) => {
    expect(isLoopback(url)).toBe(true)
  })

  it.each([
    'http://localhost.evil.com/v1',
    'https://api.example.com/v1',
    'http://example.com/localhost',
    'http://128.0.0.1/v1',
    'http://10.0.0.2/v1',
    'not an address',
  ])('treats %s as remote', (url) => {
    expect(isLoopback(url)).toBe(false)
  })
})
