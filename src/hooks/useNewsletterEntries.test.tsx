import { act, renderHook } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { NewsletterEntry } from '../lib/newsletterTypes'
import {
  createNewsletterApiEntry,
  deleteNewsletterApiEntry,
  fetchNewsletterEntries,
  fetchNewsletterEntry,
  updateNewsletterApiEntry,
  type NewsletterApiEntry,
} from '../api/newslettersApi'
import { useNewsletterEntries } from './useNewsletterEntries'

vi.mock('../api/newslettersApi', async () => {
  const actual =
    await vi.importActual<typeof import('../api/newslettersApi')>(
      '../api/newslettersApi'
    )

  return {
    ...actual,
    fetchNewsletterEntries: vi.fn(),
    fetchNewsletterEntry: vi.fn(),
    createNewsletterApiEntry: vi.fn(),
    updateNewsletterApiEntry: vi.fn(),
    deleteNewsletterApiEntry: vi.fn(),
  }
})

const mockedFetchEntries = vi.mocked(fetchNewsletterEntries)
const mockedFetchEntry = vi.mocked(fetchNewsletterEntry)
const mockedCreateEntry = vi.mocked(createNewsletterApiEntry)
const mockedUpdateEntry = vi.mocked(updateNewsletterApiEntry)
const mockedDeleteEntry = vi.mocked(deleteNewsletterApiEntry)

const sampleApiEntry: NewsletterApiEntry = {
  id: 11,
  title: 'June Digest',
  category: 'csaa',
  send_date: '2026-06-10',
  content_html: '<p>Digest</p>',
  status: 'draft',
  visibility: 'public',
  publish_at: null,
  created_at: '2026-06-01T00:00:00Z',
  updated_at: '2026-06-02T00:00:00Z',
}

const sampleLocalEntry: NewsletterEntry = {
  id: '11',
  title: 'June Digest',
  category: 'csaa',
  sendDate: '2026-06-10',
  contentHtml: '<p>Digest</p>',
  status: 'draft',
  visibility: 'public',
  publishAt: null,
  media: [],
  createdAt: '2026-06-01T00:00:00Z',
  updatedAt: '2026-06-02T00:00:00Z',
}

describe('useNewsletterEntries', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('fetches entries, maps filters, and exposes get()', async () => {
    mockedFetchEntries.mockResolvedValue({
      items: [sampleApiEntry],
      total: 1,
      page: 2,
      page_size: 5,
      total_pages: 3,
    })

    const { result } = renderHook(() => useNewsletterEntries(1, 10))

    await act(async () => {
      await result.current.fetch(2, 5, {
        status: 'draft',
        visibility: 'public',
        search: 'June',
        sortBy: 'title',
        sortOrder: 'asc',
      })
    })

    expect(mockedFetchEntries).toHaveBeenCalledWith({
      page: 2,
      page_size: 5,
      sort_by: 'title',
      sort_order: 'asc',
      status: 'draft',
      visibility: 'public',
      search: 'June',
    })
    expect(result.current.entries).toHaveLength(1)
    expect(result.current.total).toBe(1)
    expect(result.current.page).toBe(2)
    expect(result.current.pageSize).toBe(5)
    expect(result.current.totalPages).toBe(3)
    expect(result.current.get('11')?.title).toBe('June Digest')
  })

  it('uses send_date as the default sort mapping', async () => {
    mockedFetchEntries.mockResolvedValue({
      items: [sampleApiEntry],
      total: 1,
      page: 1,
      page_size: 10,
      total_pages: 1,
    })

    const { result } = renderHook(() => useNewsletterEntries(3, 25))

    await act(async () => {
      await result.current.fetch()
    })

    expect(mockedFetchEntries).toHaveBeenCalledWith({
      page: 3,
      page_size: 25,
      sort_by: 'send_date',
      sort_order: 'desc',
    })
  })

  it('surfaces fetch failures in state and rethrows them', async () => {
    const failure = new Error('Network unavailable')
    mockedFetchEntries.mockRejectedValue(failure)

    const { result } = renderHook(() => useNewsletterEntries())

    await act(async () => {
      await expect(result.current.fetch()).rejects.toThrow('Network unavailable')
    })

    expect(result.current.loading).toBe(false)
    expect(result.current.error).toBe('Network unavailable')
    expect(result.current.entries).toHaveLength(0)
  })

  it('creates entries and prepends the hydrated detail result', async () => {
    mockedCreateEntry.mockResolvedValue({
      message: 'created',
      entry: {
        id: 11,
        title: 'June Digest',
        category: 'csaa',
        send_date: '2026-06-10',
        status: 'draft',
        visibility: 'public',
      },
    })
    mockedFetchEntry.mockResolvedValue(sampleLocalEntry)

    const { result } = renderHook(() => useNewsletterEntries())

    let created: NewsletterEntry | undefined
    await act(async () => {
      created = await result.current.create({
        title: 'June Digest',
        category: 'csaa',
        sendDate: '2026-06-10',
        contentHtml: '<p>Digest</p>',
        status: 'draft',
        visibility: 'public',
        publishAt: null,
      })
    })

    expect(mockedCreateEntry).toHaveBeenCalledWith({
      title: 'June Digest',
      category: 'csaa',
      send_date: '2026-06-10',
      content_html: '<p>Digest</p>',
      status: 'draft',
      visibility: 'public',
      publish_at: null,
    })
    expect(mockedFetchEntry).toHaveBeenCalledWith('11')
    expect(created?.id).toBe('11')
    expect(result.current.entries[0]?.id).toBe('11')
    expect(result.current.total).toBe(1)
  })

  it('updates an existing entry with the refreshed detail result', async () => {
    mockedFetchEntries.mockResolvedValue({
      items: [sampleApiEntry],
      total: 1,
      page: 1,
      page_size: 10,
      total_pages: 1,
    })
    mockedUpdateEntry.mockResolvedValue({
      message: 'updated',
      entry: {
        id: 11,
        title: 'June Digest Revised',
        category: 'cst',
        send_date: '2026-06-11',
        status: 'published',
        visibility: 'private',
      },
    })
    mockedFetchEntry.mockResolvedValue({
      ...sampleLocalEntry,
      title: 'June Digest Revised',
      category: 'cst',
      sendDate: '2026-06-11',
      status: 'published',
      visibility: 'private',
    })

    const { result } = renderHook(() => useNewsletterEntries())

    await act(async () => {
      await result.current.fetch()
    })

    await act(async () => {
      await result.current.update('11', {
        title: 'June Digest Revised',
        category: 'cst',
        sendDate: '2026-06-11',
        status: 'published',
        visibility: 'private',
      })
    })

    expect(mockedUpdateEntry).toHaveBeenCalledWith('11', {
      title: 'June Digest Revised',
      category: 'cst',
      send_date: '2026-06-11',
      content_html: undefined,
      status: 'published',
      visibility: 'private',
      publish_at: undefined,
    })
    expect(result.current.entries[0]).toMatchObject({
      id: '11',
      title: 'June Digest Revised',
      category: 'cst',
      status: 'published',
    })
  })

  it('removes entries and decrements total', async () => {
    mockedFetchEntries.mockResolvedValue({
      items: [sampleApiEntry],
      total: 1,
      page: 1,
      page_size: 10,
      total_pages: 1,
    })
    mockedDeleteEntry.mockResolvedValue(undefined)

    const { result } = renderHook(() => useNewsletterEntries())

    await act(async () => {
      await result.current.fetch()
    })

    await act(async () => {
      await result.current.remove('11')
    })

    expect(mockedDeleteEntry).toHaveBeenCalledWith('11')
    expect(result.current.entries).toHaveLength(0)
    expect(result.current.total).toBe(0)
  })
})
