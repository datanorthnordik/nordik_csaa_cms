import { act, renderHook } from '@testing-library/react'
import { beforeEach, describe, expect, it } from 'vitest'
import { PRESS_ENTRIES_STORAGE_KEY } from '../lib/pressEntriesStorage'
import { usePressEntries } from './usePressEntries'
import type { PressEntry } from '../lib/pressTypes'

function seedEntry(overrides: Partial<PressEntry> = {}): PressEntry {
  return {
    id: 'seed-1',
    title: 'Seeded Entry',
    releaseDate: '2026-01-01',
    categoryId: null,
    sourceUrl: '',
    contentHtml: '',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z',
    ...overrides,
  }
}

beforeEach(() => {
  window.localStorage.clear()
  window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([seedEntry()]))
})

describe('usePressEntries', () => {
  it('returns the initial entries from localStorage', () => {
    const { result } = renderHook(() => usePressEntries())
    expect(result.current.entries).toHaveLength(1)
    expect(result.current.entries[0].id).toBe('seed-1')
  })

  it('create adds a new entry and updates the entries list', () => {
    const { result } = renderHook(() => usePressEntries())
    act(() => {
      result.current.create({
        title: 'Created',
        releaseDate: '2026-06-01',
        categoryId: null,
        sourceUrl: '',
        contentHtml: '',
        status: 'draft',
        visibility: 'public',
        publishAt: null,
        media: [],
      })
    })
    expect(result.current.entries).toHaveLength(2)
    expect(result.current.entries.some((e) => e.title === 'Created')).toBe(true)
  })

  it('update patches an existing entry', () => {
    const { result } = renderHook(() => usePressEntries())
    act(() => {
      result.current.update('seed-1', { title: 'Updated Title' })
    })
    expect(result.current.entries[0].title).toBe('Updated Title')
  })

  it('remove deletes an entry and reflects in entries list', () => {
    const { result } = renderHook(() => usePressEntries())
    act(() => {
      result.current.remove('seed-1')
    })
    expect(result.current.entries.find((e) => e.id === 'seed-1')).toBeUndefined()
  })

  it('get returns the entry with the matching id', () => {
    const { result } = renderHook(() => usePressEntries())
    const found = result.current.get('seed-1')
    expect(found?.title).toBe('Seeded Entry')
  })

  it('get returns undefined for an unknown id', () => {
    const { result } = renderHook(() => usePressEntries())
    expect(result.current.get('no-such-id')).toBeUndefined()
  })

  it('refresh re-reads entries from localStorage', () => {
    const { result } = renderHook(() => usePressEntries())
    const extra = seedEntry({ id: 'extra', title: 'Extra' })
    const current = JSON.parse(
      window.localStorage.getItem(PRESS_ENTRIES_STORAGE_KEY) ?? '[]',
    ) as PressEntry[]
    window.localStorage.setItem(
      PRESS_ENTRIES_STORAGE_KEY,
      JSON.stringify([extra, ...current]),
    )
    act(() => {
      result.current.refresh()
    })
    expect(result.current.entries.some((e) => e.id === 'extra')).toBe(true)
  })
})
