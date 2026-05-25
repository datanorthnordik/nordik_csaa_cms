import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import { memorialApi } from './memorialApi'
import type { MemorialListFilters } from '../lib/memorialTypes'

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

const defaultFilters: MemorialListFilters = {
  page: 1,
  pageSize: 10,
  searchTerm: '',
  status: 'all',
  category: '',
}

const sampleListResponse = {
  data: {
    items: [
      {
        id: 12,
        full_name: 'Ada Lovelace',
        affiliation: 'Analytical Engine',
        category: 'founder' as const,
        category_label: 'Founder',
        status: 'published' as const,
        date_of_birth: '1815-12-10',
        date_of_passing: '1852-11-27',
        created_at: '2026-05-20T00:00:00Z',
        updated_at: '2026-05-21T00:00:00Z',
        published_at: '2026-05-21T00:00:00Z',
      },
    ],
    pagination: {
      page: 1,
      page_size: 10,
      total_items: 1,
      total_pages: 1,
      has_next: false,
      has_prev: false,
    },
    applied_filters: {
      page: 1,
      page_size: 10,
      search_term: '',
      status: 'all',
      category: '',
    },
  },
}

beforeEach(() => {
  vi.clearAllMocks()
})

describe('memorialApi', () => {
  describe('listMemorials', () => {
    it('requests the memorial endpoint with normalized query params and maps the response', async () => {
      mockedGet.mockResolvedValue(sampleListResponse)

      const result = await memorialApi.listMemorials({
        ...defaultFilters,
        searchTerm: '  Ada  ',
        status: 'published',
        category: 'founder',
      })

      expect(mockedGet).toHaveBeenCalledOnce()
      const [url, config] = mockedGet.mock.calls[0]
      const params = config?.params as URLSearchParams
      expect(String(url)).toContain('/api/memorial')
      expect(params.get('page')).toBe('1')
      expect(params.get('page_size')).toBe('10')
      expect(params.get('status')).toBe('published')
      expect(params.get('search')).toBe('Ada')
      expect(params.get('category')).toBe('founder')

      expect(result).toEqual({
        items: [
          {
            id: '12',
            fullName: 'Ada Lovelace',
            affiliation: 'Analytical Engine',
            category: 'founder',
            categoryLabel: 'Founder',
            status: 'published',
            dateOfBirth: '1815-12-10',
            dateOfPassing: '1852-11-27',
            createdAt: '2026-05-20T00:00:00Z',
            updatedAt: '2026-05-21T00:00:00Z',
            publishedAt: '2026-05-21T00:00:00Z',
          },
        ],
        pagination: {
          page: 1,
          pageSize: 10,
          totalItems: 1,
          totalPages: 1,
          hasNext: false,
          hasPrev: false,
        },
      })
    })

    it('omits optional search and category params when they are blank', async () => {
      mockedGet.mockResolvedValue(sampleListResponse)

      await memorialApi.listMemorials(defaultFilters)

      const [, config] = mockedGet.mock.calls[0]
      const params = config?.params as URLSearchParams
      expect(params.get('status')).toBe('all')
      expect(params.has('search')).toBe(false)
      expect(params.has('category')).toBe(false)
    })
  })

  describe('getMemorial', () => {
    it('maps memorial detail responses into the local shape', async () => {
      mockedGet.mockResolvedValue({
        data: {
          id: 18,
          full_name: 'Grace Hopper',
          affiliation: 'US Navy',
          category: 'veteran',
          category_label: 'Veteran',
          status: 'review',
          biography: '<p>Pioneer</p>',
          date_of_birth: '1906-12-09',
          date_of_passing: '1992-01-01',
          created_at: '2026-05-18T00:00:00Z',
          updated_at: '2026-05-19T00:00:00Z',
          portrait: {
            file_name: 'portrait.jpg',
            mime_type: 'image/jpeg',
            file_size: 2048,
          },
          gallery_images: [
            {
              id: 91,
              file_name: 'gallery-one.png',
              mime_type: 'image/png',
              file_size: 4096,
            },
          ],
        },
      })

      const result = await memorialApi.getMemorial('18')

      expect(mockedGet).toHaveBeenCalledWith('/api/memorial/18')
      expect(result).toEqual({
        id: '18',
        fullName: 'Grace Hopper',
        affiliation: 'US Navy',
        category: 'veteran',
        categoryLabel: 'Veteran',
        status: 'review',
        biography: '<p>Pioneer</p>',
        dateOfBirth: '1906-12-09',
        dateOfPassing: '1992-01-01',
        createdAt: '2026-05-18T00:00:00Z',
        updatedAt: '2026-05-19T00:00:00Z',
        publishedAt: undefined,
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
      })
    })
  })

  describe('media content requests', () => {
    it('requests portrait content as a blob without showing the shared error toast', async () => {
      const blob = new Blob(['portrait'], { type: 'image/jpeg' })
      mockedGet.mockResolvedValue({ data: blob })

      const result = await memorialApi.getMemorialPortraitContent('4')

      expect(result).toBe(blob)
      expect(mockedGet).toHaveBeenCalledWith('/api/memorial/4/portrait/content', {
        responseType: 'blob',
        skipErrorToast: true,
      })
    })

    it('requests gallery image content as a blob without showing the shared error toast', async () => {
      const blob = new Blob(['gallery'], { type: 'image/png' })
      mockedGet.mockResolvedValue({ data: blob })

      const result = await memorialApi.getMemorialGalleryImageContent('4', '7')

      expect(result).toBe(blob)
      expect(mockedGet).toHaveBeenCalledWith('/api/memorial/4/gallery/7/content', {
        responseType: 'blob',
        skipErrorToast: true,
      })
    })
  })

  describe('createMemorial', () => {
    it('posts a plain object when there are no uploaded files', async () => {
      mockedPost.mockResolvedValue({
        data: {
          message: 'created',
          memorial: {
            id: 7,
            full_name: 'Ada Lovelace',
            category: 'founder',
            status: 'draft',
            updated_at: '2026-05-20T00:00:00Z',
          },
        },
      })

      const result = await memorialApi.createMemorial({
        full_name: 'Ada Lovelace',
        affiliation: 'Analytical Engine',
        category: 'founder',
        status: 'draft',
        biography: '<p>Remembered</p>',
      })

      expect(result.id).toBe(7)
      const [, body] = mockedPost.mock.calls[0]
      expect(body instanceof FormData).toBe(false)
      expect(body).toMatchObject({
        full_name: 'Ada Lovelace',
        affiliation: 'Analytical Engine',
        category: 'founder',
        status: 'draft',
        biography: '<p>Remembered</p>',
        date_of_birth: '',
        date_of_passing: '',
        remove_portrait: false,
        remove_gallery_image_ids: [],
        gallery_images: [],
      })
    })

    it('posts FormData with portrait and gallery file metadata when uploads are included', async () => {
      mockedPost.mockResolvedValue({
        data: {
          message: 'created',
          memorial: {
            id: 8,
            full_name: 'Grace Hopper',
            category: 'veteran',
            status: 'published',
            updated_at: '2026-05-20T00:00:00Z',
          },
        },
      })

      const portraitFile = new File(['portrait-bytes'], 'portrait.jpg', {
        type: 'image/jpeg',
      })
      const galleryOne = new File(['gallery-one'], 'gallery-one.png', {
        type: 'image/png',
      })
      const galleryTwo = new File(['gallery-two'], 'gallery-two.webp', {
        type: 'image/webp',
      })

      await memorialApi.createMemorial(
        {
          full_name: 'Grace Hopper',
          affiliation: 'US Navy',
          category: 'veteran',
          status: 'published',
          biography: '<p>Pioneer</p>',
          date_of_birth: '1906-12-09',
          date_of_passing: '1992-01-01',
          remove_portrait: true,
          remove_gallery_image_ids: [31, 32],
        },
        portraitFile,
        [galleryOne, galleryTwo],
      )

      const [, body] = mockedPost.mock.calls[0]
      expect(body instanceof FormData).toBe(true)

      const payload = JSON.parse((body as FormData).get('payload') as string)
      expect(payload).toEqual({
        full_name: 'Grace Hopper',
        affiliation: 'US Navy',
        category: 'veteran',
        status: 'published',
        biography: '<p>Pioneer</p>',
        date_of_birth: '1906-12-09',
        date_of_passing: '1992-01-01',
        remove_portrait: true,
        remove_gallery_image_ids: [31, 32],
        portrait: {
          file_name: 'portrait.jpg',
          mime_type: 'image/jpeg',
          file_size: portraitFile.size,
        },
        gallery_images: [
          {
            file_name: 'gallery-one.png',
            mime_type: 'image/png',
            file_size: galleryOne.size,
          },
          {
            file_name: 'gallery-two.webp',
            mime_type: 'image/webp',
            file_size: galleryTwo.size,
          },
        ],
      })
      expect(((body as FormData).get('portrait_file') as File).name).toBe(
        portraitFile.name,
      )
      expect(
        ((body as FormData).get('gallery_images[0].file') as File).name,
      ).toBe(galleryOne.name)
      expect(
        ((body as FormData).get('gallery_images[1].file') as File).name,
      ).toBe(galleryTwo.name)
    })
  })

  describe('updateMemorial', () => {
    it('puts to the memorial by id endpoint', async () => {
      mockedPut.mockResolvedValue({
        data: {
          message: 'updated',
          memorial: {
            id: 9,
            full_name: 'Ada Updated',
            category: 'friend',
            status: 'review',
            updated_at: '2026-05-21T00:00:00Z',
          },
        },
      })

      await memorialApi.updateMemorial('9', {
        full_name: 'Ada Updated',
        affiliation: 'CSAA',
        category: 'friend',
        status: 'review',
        biography: '<p>Updated</p>',
      })

      const [url] = mockedPut.mock.calls[0]
      expect(String(url)).toContain('/api/memorial/9')
    })
  })

  describe('deleteMemorial', () => {
    it('sends a delete request to the memorial by id endpoint', async () => {
      mockedDelete.mockResolvedValue({})

      await memorialApi.deleteMemorial('44')

      expect(mockedDelete).toHaveBeenCalledWith('/api/memorial/44')
    })
  })
})
