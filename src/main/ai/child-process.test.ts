import { EventEmitter } from 'node:events'
import type { ChildProcessWithoutNullStreams } from 'node:child_process'
import { describe, expect, it, vi } from 'vitest'
import { killTree, terminateNow, type ProcessTreeKiller } from './child-process.js'

function fakeChild(pid: number | undefined): ChildProcessWithoutNullStreams & { signals: string[] } {
  const signals: string[] = []
  const child = Object.assign(new EventEmitter(), {
    pid,
    exitCode: null,
    signalCode: null,
    signals,
    kill: (signal: string) => signals.push(signal) > 0,
  })
  return child as unknown as ChildProcessWithoutNullStreams & { signals: string[] }
}

function killer(windows: boolean): ProcessTreeKiller & { groups: unknown[]; taskkills: unknown[] } {
  const groups: unknown[] = []
  const taskkills: unknown[] = []
  return {
    windows,
    groups,
    taskkills,
    killGroup: (pid, signal) => groups.push([pid, signal]),
    runTaskkill: (args) => taskkills.push(args),
  }
}

describe('killTree', () => {
  it('on POSIX sends the signal to the whole program group', () => {
    const tree = killer(false)
    killTree(fakeChild(42), 'SIGTERM', tree)
    expect(tree.groups).toEqual([[42, 'SIGTERM']])
  })

  it('on Windows ends the tree with taskkill', () => {
    const tree = killer(true)
    killTree(fakeChild(42), 'SIGTERM', tree)
    expect(tree.taskkills).toEqual([['/pid', '42', '/T', '/F']])
  })

  it('without a group, ends only the program', () => {
    const child = fakeChild(42)
    const tree: ProcessTreeKiller = {
      ...killer(false),
      killGroup: () => {
        throw new Error('ESRCH')
      },
    }
    killTree(child, 'SIGTERM', tree)
    expect(child.signals).toEqual(['SIGTERM'])
  })
})

describe('terminateNow', () => {
  it('on exit, kills with SIGKILL immediately, without waiting for a timer', () => {
    vi.useFakeTimers()
    const tree = killer(false)
    terminateNow(fakeChild(7), tree)
    expect(tree.groups).toEqual([[7, 'SIGKILL']])
    expect(vi.getTimerCount()).toBe(0)
    vi.useRealTimers()
  })
})
