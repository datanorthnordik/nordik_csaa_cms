import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import { eventsApi } from './eventsApi'
import type { EventListFilters } from './eventsApi'

vi.mock('./apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)
const mockedPost = vi.mocked(apiClient.post)
const mockedPut = vi.mocked(apiClient.put)
const mockedDelete = vi.mocked(apiClient.delete)

const defaultFilters: EventListFilters = {
  page: 1,
  pageSize: 10,
  searchTerm: '',
  statuses: [],
  startDate: '',
  endDate: '',
  dateRange: 'custom',
  sortBy: 'start_at',
  sortOrder: 'desc',
}

const emptyListResponse = {
  data: {
    items: [],
    pagination: { page: 1, page_size: 10, total_items: 0, total_pages: 0, has_next: false, has_prev: false },
    applied_filters: {
      page: 1, page_size: 10, search_term: '', statuses: [],
      date_range: 'custom', sort_by: 'start_at', sort_order: 'desc',
    },
  },
}

beforeEach(() => {
  vi.clearAllMocks()
})

describe('eventsApi', () => {
  describe('listEvents', () => {
    it('requests the events endpoint with page and page_size params', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      await eventsApi.listEvents(defaultFilters)
      expect(mockedGet).toHaveBeenCalledOnce()
      const [, config] = mockedGet.mock.calls[0]
      const params = config?.params as URLSearchParams
      expect(params.get('page')).toBe('1')
      expect(params.get('page_size')).toBe('10')
    })

    it('includes search param when searchTerm is non-empty', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      await eventsApi.listEvents({ ...defaultFilters, searchTerm: '  Spring  ' })
      const [, config] = mockedGet.mock.calls[0]
      expect((config?.params as URLSearchParams).get('search')).toBe('Spring')
    })

    it('omits search param when searchTerm is blank', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      await eventsApi.listEvents(defaultFilters)
      const [, config] = mockedGet.mock.calls[0]
      expect((config?.params as URLSearchParams).has('search')).toBe(false)
    })

    it('appends a status param for each status in the filter', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      await eventsApi.listEvents({ ...defaultFilters, statuses: ['published', 'draft'] })
      const [, config] = mockedGet.mock.calls[0]
      expect((config?.params as URLSearchParams).getAll('status')).toEqual(['published', 'draft'])
    })

    it('includes start_date and end_date when set', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      await eventsApi.listEvents({ ...defaultFilters, startDate: '2026-01-01', endDate: '2026-12-31' })
      const [, config] = mockedGet.mock.calls[0]
      const params = config?.params as URLSearchParams
      expect(params.get('start_date')).toBe('2026-01-01')
      expect(params.get('end_date')).toBe('2026-12-31')
    })

    it('returns the response data', async () => {
      mockedGet.mockResolvedValue(emptyListResponse)
      const result = await eventsApi.listEvents(defaultFilters)
      expect(result.items).toEqual([])
    })
  })

  describe('getEvent', () => {
    it('requests the event by id and returns data', async () => {
      mockedGet.mockResolvedValue({ data: { id: 42, title: 'Annual Feast' } })
      const result = await eventsApi.getEvent(42)
      expect(result.id).toBe(42)
      const [url] = mockedGet.mock.calls[0]
      expect(String(url)).toContain('42')
    })
  })

  describe('createEvent', () => {
    it('posts a plain object when no files are attached', async () => {
      mockedPost.mockResolvedValue({ data: { message: 'ok', event: { id: 1, title: 'New', published: false } } })
      await eventsApi.createEvent({ title: 'New', attachments: [], occurrences: [] } as any)
      const [, body] = mockedPost.mock.calls[0]
      expect(body instanceof FormData).toBe(false)
    })

    it('posts FormData when a display image file is included', async () => {
      mockedPost.mockResolvedValue({ data: { message: 'ok', event: { id: 1, title: 'New', published: false } } })
      const file = new File(['img'], 'photo.jpg', { type: 'image/jpeg' })
      await eventsApi.createEvent({
        title: 'New',
        attachments: [],
        occurrences: [],
        displayImageFile: file,
      } as any)
      const [, body] = mockedPost.mock.calls[0]
      expect(body instanceof FormData).toBe(true)
    })

    it('posts FormData when attachment files are included', async () => {
      mockedPost.mockResolvedValue({ data: { message: 'ok', event: { id: 1, title: 'New', published: false } } })
      const file = new File(['doc'], 'doc.pdf', { type: 'application/pdf' })
      await eventsApi.createEvent({
        title: 'New',
        attachments: [{}],
        occurrences: [],
        attachmentFiles: [file],
      } as any)
      const [, body] = mockedPost.mock.calls[0]
      expect(body instanceof FormData).toBe(true)
    })
  })

  describe('updateEvent', () => {
    it('puts to the event by id endpoint', async () => {
      mockedPut.mockResolvedValue({ data: { message: 'ok', event: { id: 7, title: 'Updated', published: true } } })
      await eventsApi.updateEvent(7, { title: 'Updated', attachments: [], occurrences: [] } as any)
      const [url] = mockedPut.mock.calls[0]
      expect(String(url)).toContain('7')
    })
  })

  describe('deleteEvent', () => {
    it('sends a delete request to the event by id endpoint', async () => {
      mockedDelete.mockResolvedValue({})
      await eventsApi.deleteEvent(99)
      const [url] = mockedDelete.mock.calls[0]
      expect(String(url)).toContain('99')
    })
  })

  describe('deleteEventDocument', () => {
    it('passes storage_url as a query parameter', async () => {
      mockedDelete.mockResolvedValue({})
      await eventsApi.deleteEventDocument(5, 'gs://bucket/file.pdf')
      const [, config] = mockedDelete.mock.calls[0]
      expect((config as any)?.params?.storage_url).toBe('gs://bucket/file.pdf')
    })
  })

  describe('deleteEventPhoto', () => {
    it('passes storage_url as a query parameter', async () => {
      mockedDelete.mockResolvedValue({})
      await eventsApi.deleteEventPhoto(3, 'gs://bucket/photo.jpg')
      const [, config] = mockedDelete.mock.calls[0]
      expect((config as any)?.params?.storage_url).toBe('gs://bucket/photo.jpg')
    })
  })
})
