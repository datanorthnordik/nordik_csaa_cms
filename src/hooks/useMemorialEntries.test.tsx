import { act, renderHook } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { memorialApi } from '../api/memorialApi'
import type { MemorialEntry, MemorialEntrySummary } from '../lib/memorialTypes'
import { useMemorialEntries } from './useMemorialEntries'

vi.mock('../api/memorialApi', () => ({
  memorialApi: {
    listMemorials: vi.fn(),
    getMemorial: vi.fn(),
    createMemorial: vi.fn(),
    updateMemorial: vi.fn(),
    deleteMemorial: vi.fn(),
  },
}))

const mockedListMemorials = vi.mocked(memorialApi.listMemorials)
const mockedGetMemorial = vi.mocked(memorialApi.getMemorial)
const mockedCreateMemorial = vi.mocked(memorialApi.createMemorial)
const mockedUpdateMemorial = vi.mocked(memorialApi.updateMemorial)
const mockedDeleteMemorial = vi.mocked(memorialApi.deleteMemorial)

const sampleSummary: MemorialEntrySummary = {
  id: '11',
  fullName: 'Ada Lovelace',
  affiliation: 'Analytical Engine',
  category: 'founder',
  categoryLabel: 'Founder',
  status: 'draft',
  dateOfBirth: '1815-12-10',
  dateOfPassing: '1852-11-27',
  createdAt: '2026-05-20T00:00:00Z',
  updatedAt: '2026-05-21T00:00:00Z',
  publishedAt: undefined,
}

const sampleEntry: MemorialEntry = {
  ...sampleSummary,
  biography: '<p>Remembered</p>',
  portrait: {
    fileName: 'portrait.jpg',
    mimeType: 'image/jpeg',
    fileSize: 2048,
  },
  galleryImages: [
    {
      id: '91',
      fileName: 'gallery-one.png',
      mimeType: 'image/png',
      fileSize: 4096,
    },
  ],
}

describe('useMemorialEntries', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('fetches memorial entries and updates loading, items, and pagination state', async () => {
    mockedListMemorials.mockResolvedValue({
      items: [sampleSummary],
      pagination: {
        page: 2,
        pageSize: 5,
        totalItems: 8,
        totalPages: 2,
        hasNext: false,
        hasPrev: true,
      },
    })

    const { result } = renderHook(() => useMemorialEntries(5))

    await act(async () => {
      await result.current.fetch({
        searchTerm: 'Ada',
        status: 'published',
        category: 'founder',
        page: 2,
        pageSize: 5,
      })
    })

    expect(mockedListMemorials).toHaveBeenCalledWith({
      searchTerm: 'Ada',
      status: 'published',
      category: 'founder',
      page: 2,
      pageSize: 5,
    })
    expect(result.current.items).toEqual([sampleSummary])
    expect(result.current.pagination).toEqual({
      page: 2,
      pageSize: 5,
      totalItems: 8,
      totalPages: 2,
      hasNext: false,
      hasPrev: true,
    })
    expect(result.current.loading).toBe(false)
    expect(result.current.error).toBeNull()
  })

  it('surfaces fetch failures in state and uses the fallback message for non-Error rejections', async () => {
    mockedListMemorials.mockRejectedValue('network unavailable')

    const { result } = renderHook(() => useMemorialEntries())

    await act(async () => {
      await expect(
        result.current.fetch({
          searchTerm: '',
          status: 'all',
          category: '',
          page: 1,
          pageSize: 10,
        }),
      ).rejects.toBe('network unavailable')
    })

    expect(result.current.loading).toBe(false)
    expect(result.current.error).toBe('Failed to fetch memorial entries')
    expect(result.current.items).toEqual([])
  })

  it('creates a memorial entry and hydrates the full detail response', async () => {
    const portraitFile = new File(['portrait'], 'portrait.jpg', { type: 'image/jpeg' })
    const galleryFile = new File(['gallery'], 'gallery.png', { type: 'image/png' })

    mockedCreateMemorial.mockResolvedValue({
      id: 11,
      full_name: 'Ada Lovelace',
      category: 'founder',
      status: 'draft',
      updated_at: '2026-05-21T00:00:00Z',
    } as never)
    mockedGetMemorial.mockResolvedValue(sampleEntry)

    const { result } = renderHook(() => useMemorialEntries())

    let created: MemorialEntry | undefined
    await act(async () => {
      created = await result.current.create(
        {
          full_name: 'Ada Lovelace',
          affiliation: 'Analytical Engine',
          category: 'founder',
          status: 'draft',
          biography: '<p>Remembered</p>',
        },
        portraitFile,
        [galleryFile],
      )
    })

    expect(mockedCreateMemorial).toHaveBeenCalledWith(
      {
        full_name: 'Ada Lovelace',
        affiliation: 'Analytical Engine',
        category: 'founder',
        status: 'draft',
        biography: '<p>Remembered</p>',
      },
      portraitFile,
      [galleryFile],
    )
    expect(mockedGetMemorial).toHaveBeenCalledWith('11')
    expect(created).toEqual(sampleEntry)
  })

  it('updates a memorial entry and rehydrates the latest detail response', async () => {
    mockedUpdateMemorial.mockResolvedValue({
      id: 11,
      full_name: 'Ada Updated',
      category: 'friend',
      status: 'review',
      updated_at: '2026-05-22T00:00:00Z',
    } as never)
    mockedGetMemorial.mockResolvedValue({
      ...sampleEntry,
      fullName: 'Ada Updated',
      category: 'friend',
      categoryLabel: 'Friend',
      status: 'review',
    })

    const { result } = renderHook(() => useMemorialEntries())

    let updated: MemorialEntry | undefined
    await act(async () => {
      updated = await result.current.update('11', {
        full_name: 'Ada Updated',
        affiliation: 'CSAA',
        category: 'friend',
        status: 'review',
        biography: '<p>Updated</p>',
      })
    })

    expect(mockedUpdateMemorial).toHaveBeenCalledWith('11', {
      full_name: 'Ada Updated',
      affiliation: 'CSAA',
      category: 'friend',
      status: 'review',
      biography: '<p>Updated</p>',
    }, undefined, [])
    expect(mockedGetMemorial).toHaveBeenCalledWith('11')
    expect(updated?.fullName).toBe('Ada Updated')
  })

  it('removes memorial entries from local state and decrements total items without going negative', async () => {
    mockedListMemorials.mockResolvedValue({
      items: [sampleSummary],
      pagination: {
        page: 1,
        pageSize: 10,
        totalItems: 1,
        totalPages: 1,
        hasNext: false,
        hasPrev: false,
      },
    })
    mockedDeleteMemorial.mockResolvedValue(undefined)

    const { result } = renderHook(() => useMemorialEntries())

    await act(async () => {
      await result.current.fetch({
        searchTerm: '',
        status: 'all',
        category: '',
        page: 1,
        pageSize: 10,
      })
    })

    await act(async () => {
      await result.current.remove('11')
    })

    expect(mockedDeleteMemorial).toHaveBeenCalledWith('11')
    expect(result.current.items).toEqual([])
    expect(result.current.pagination.totalItems).toBe(0)
  })
})
