import { beforeEach, describe, expect, it } from 'vitest'
import {
  createPressEntry,
  deletePressEntry,
  getPressEntry,
  loadPressEntries,
  PRESS_ENTRIES_STORAGE_KEY,
  savePressEntries,
  updatePressEntry,
} from './pressEntriesStorage'
import type { PressEntry } from './pressTypes'

function emptyStorage() {
  window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([]))
}

function makeEntry(overrides: Partial<PressEntry> = {}): PressEntry {
  return {
    id: 'test-id',
    title: 'Test Entry',
    releaseDate: '2026-01-15',
    categoryId: null,
    sourceUrl: 'https://example.com',
    contentHtml: '<p>Hello</p>',
    status: 'draft',
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
})

describe('pressEntriesStorage', () => {
  describe('loadPressEntries', () => {
    it('seeds with default data when storage is empty', () => {
      const entries = loadPressEntries()
      expect(entries.length).toBeGreaterThan(0)
    })

    it('returns stored entries when storage has valid data', () => {
      const entry = makeEntry()
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([entry]))
      const entries = loadPressEntries()
      expect(entries).toHaveLength(1)
      expect(entries[0].id).toBe('test-id')
    })

    it('falls back to seed data when storage contains invalid JSON', () => {
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, 'not-valid-json')
      const entries = loadPressEntries()
      expect(entries.length).toBeGreaterThan(0)
    })

    it('filters out entries that are missing required fields', () => {
      const invalid = [{ id: 'x' }, makeEntry({ id: 'valid' })]
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify(invalid))
      const entries = loadPressEntries()
      expect(entries.every((e) => e.id === 'valid')).toBe(true)
    })
  })

  describe('savePressEntries', () => {
    it('persists entries to localStorage', () => {
      const entry = makeEntry()
      savePressEntries([entry])
      const raw = window.localStorage.getItem(PRESS_ENTRIES_STORAGE_KEY)
      expect(JSON.parse(raw!)[0].id).toBe('test-id')
    })
  })

  describe('createPressEntry', () => {
    it('prepends the new entry and generates an id and timestamps', () => {
      emptyStorage()
      const created = createPressEntry({
        title: 'New',
        releaseDate: '2026-06-01',
        categoryId: null,
        sourceUrl: '',
        contentHtml: '',
        status: 'draft',
        visibility: 'public',
        publishAt: null,
        media: [],
      })
      expect(typeof created.id).toBe('string')
      expect(created.id.length).toBeGreaterThan(0)
      expect(created.createdAt).toBeDefined()
      expect(created.updatedAt).toBeDefined()
      const stored = loadPressEntries()
      expect(stored[0].id).toBe(created.id)
    })

    it('new entry appears before any existing ones', () => {
      const existing = makeEntry({ id: 'old' })
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([existing]))
      const created = createPressEntry({
        title: 'Latest',
        releaseDate: '2026-06-01',
        categoryId: null,
        sourceUrl: '',
        contentHtml: '',
        status: 'draft',
        visibility: 'public',
        publishAt: null,
        media: [],
      })
      const stored = loadPressEntries()
      expect(stored[0].id).toBe(created.id)
      expect(stored[1].id).toBe('old')
    })
  })

  describe('updatePressEntry', () => {
    it('updates the patched fields while preserving id and createdAt', () => {
      const entry = makeEntry({ id: 'u1' })
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([entry]))
      const updated = updatePressEntry('u1', { title: 'Updated', status: 'published' })
      expect(updated?.title).toBe('Updated')
      expect(updated?.status).toBe('published')
      expect(updated?.id).toBe('u1')
      expect(updated?.createdAt).toBe(entry.createdAt)
    })

    it('returns undefined for a non-existent id', () => {
      emptyStorage()
      const result = updatePressEntry('missing', { title: 'X' })
      expect(result).toBeUndefined()
    })

    it('persists the update to localStorage', () => {
      const entry = makeEntry({ id: 'u2' })
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([entry]))
      updatePressEntry('u2', { title: 'Persisted' })
      const stored = loadPressEntries()
      expect(stored.find((e) => e.id === 'u2')?.title).toBe('Persisted')
    })
  })

  describe('deletePressEntry', () => {
    it('removes the entry with the given id', () => {
      const a = makeEntry({ id: 'a' })
      const b = makeEntry({ id: 'b', title: 'Keep me' })
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([a, b]))
      deletePressEntry('a')
      const stored = loadPressEntries()
      expect(stored.find((e) => e.id === 'a')).toBeUndefined()
      expect(stored.find((e) => e.id === 'b')).toBeDefined()
    })

    it('is a no-op for a non-existent id', () => {
      const entry = makeEntry()
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([entry]))
      deletePressEntry('does-not-exist')
      expect(loadPressEntries()).toHaveLength(1)
    })
  })

  describe('getPressEntry', () => {
    it('returns the matching entry', () => {
      const entry = makeEntry({ id: 'find-me', title: 'Findable' })
      window.localStorage.setItem(PRESS_ENTRIES_STORAGE_KEY, JSON.stringify([entry]))
      expect(getPressEntry('find-me')?.title).toBe('Findable')
    })

    it('returns undefined when the id does not exist', () => {
      emptyStorage()
      expect(getPressEntry('nope')).toBeUndefined()
    })
  })
})
