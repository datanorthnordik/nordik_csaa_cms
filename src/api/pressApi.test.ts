import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import {
  createPressApiEntry,
  deletePressApiEntry,
  fetchPressEntries,
  fetchPressEntry,
  localEntryToPressApiInput,
  pressApiEntryToLocal,
  updatePressApiEntry,
} from './pressApi'
import type { PressApiEntry } from './pressApi'

vi.mock('./apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)
const mockedPost = vi.mocked(apiClient.post)
const mockedPatch = vi.mocked(apiClient.patch)
const mockedDelete = vi.mocked(apiClient.delete)

const sampleApiEntry: PressApiEntry = {
  id: 'abc-123',
  title: 'Spring Release',
  release_date: '2026-03-20',
  category_id: 'cat-1',
  source_url: 'https://example.com/news',
  content_html: '<p>Body</p>',
  status: 'published',
  visibility: 'public',
  publish_at: null,
  created_at: '2026-01-01T00:00:00.000Z',
  updated_at: '2026-01-02T00:00:00.000Z',
}

beforeEach(() => {
  vi.clearAllMocks()
})

describe('pressApi converters', () => {
  describe('pressApiEntryToLocal', () => {
    it('maps snake_case fields to camelCase', () => {
      const local = pressApiEntryToLocal(sampleApiEntry)
      expect(local.id).toBe('abc-123')
      expect(local.releaseDate).toBe('2026-03-20')
      expect(local.categoryId).toBe('cat-1')
      expect(local.sourceUrl).toBe('https://example.com/news')
      expect(local.contentHtml).toBe('<p>Body</p>')
      expect(local.createdAt).toBe(sampleApiEntry.created_at)
      expect(local.updatedAt).toBe(sampleApiEntry.updated_at)
    })

    it('initialises media as an empty array', () => {
      const local = pressApiEntryToLocal(sampleApiEntry)
      expect(local.media).toEqual([])
    })
  })

  describe('localEntryToPressApiInput', () => {
    it('maps camelCase fields back to snake_case', () => {
      const local = pressApiEntryToLocal(sampleApiEntry)
      const input = localEntryToPressApiInput(local)
      expect(input.release_date).toBe('2026-03-20')
      expect(input.category_id).toBe('cat-1')
      expect(input.source_url).toBe('https://example.com/news')
      expect(input.content_html).toBe('<p>Body</p>')
    })

    it('round-trips through toLocal then toInput without data loss', () => {
      const local = pressApiEntryToLocal(sampleApiEntry)
      const input = localEntryToPressApiInput(local)
      expect(input.title).toBe(sampleApiEntry.title)
      expect(input.status).toBe(sampleApiEntry.status)
      expect(input.visibility).toBe(sampleApiEntry.visibility)
    })
  })
})

describe('pressApi CRUD', () => {
  describe('fetchPressEntries', () => {
    it('makes a GET request and returns the list response', async () => {
      const responseData = { items: [sampleApiEntry], total: 1, page: 1, page_size: 10 }
      mockedGet.mockResolvedValue({ data: responseData })
      const result = await fetchPressEntries()
      expect(mockedGet).toHaveBeenCalledOnce()
      expect(result.total).toBe(1)
    })

    it('forwards optional query params to the request', async () => {
      mockedGet.mockResolvedValue({ data: { items: [], total: 0, page: 1, page_size: 10 } })
      await fetchPressEntries({ page: 2, page_size: 5 })
      const [, config] = mockedGet.mock.calls[0]
      expect((config as any)?.params?.page).toBe(2)
    })
  })

  describe('fetchPressEntry', () => {
    it('makes a GET request for the entry by id', async () => {
      mockedGet.mockResolvedValue({ data: sampleApiEntry })
      const result = await fetchPressEntry('abc-123')
      expect(result.id).toBe('abc-123')
      const [url] = mockedGet.mock.calls[0]
      expect(String(url)).toContain('abc-123')
    })
  })

  describe('createPressApiEntry', () => {
    it('POSTs to the press endpoint with the entry payload', async () => {
      mockedPost.mockResolvedValue({ data: sampleApiEntry })
      const { id, created_at, updated_at, ...input } = sampleApiEntry
      await createPressApiEntry(input)
      expect(mockedPost).toHaveBeenCalledOnce()
      const [, body] = mockedPost.mock.calls[0]
      expect((body as PressApiEntry).title).toBe('Spring Release')
    })

    it('returns the created entry from the response', async () => {
      mockedPost.mockResolvedValue({ data: sampleApiEntry })
      const { id, created_at, updated_at, ...input } = sampleApiEntry
      const result = await createPressApiEntry(input)
      expect(result.id).toBe('abc-123')
    })
  })

  describe('updatePressApiEntry', () => {
    it('PATCHes the entry by id', async () => {
      const updated = { ...sampleApiEntry, title: 'New Title' }
      mockedPatch.mockResolvedValue({ data: updated })
      const result = await updatePressApiEntry('abc-123', { title: 'New Title' })
      expect(result.title).toBe('New Title')
      const [url] = mockedPatch.mock.calls[0]
      expect(String(url)).toContain('abc-123')
    })
  })

  describe('deletePressApiEntry', () => {
    it('sends a DELETE request to the entry by id endpoint', async () => {
      mockedDelete.mockResolvedValue({})
      await deletePressApiEntry('abc-123')
      const [url] = mockedDelete.mock.calls[0]
      expect(String(url)).toContain('abc-123')
    })
  })
})
