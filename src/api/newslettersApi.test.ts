import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import {
  addNewsletterMedia,
  createNewsletterApiEntry,
  deleteNewsletterApiEntry,
  deleteNewsletterMedia,
  fetchNewsletterEntries,
  fetchNewsletterEntry,
  getNewsletterMediaContent,
  newsletterApiDetailToLocal,
  newsletterApiEntryToLocal,
  newsletterApiMediaToLocal,
  reorderNewsletterMedia,
  updateNewsletterApiEntry,
  updateNewsletterMedia,
  type NewsletterApiDetailEntry,
  type NewsletterApiEntry,
  type NewsletterApiMedia,
} from './newslettersApi'

vi.mock('./apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)
const mockedPost = vi.mocked(apiClient.post)
const mockedPut = vi.mocked(apiClient.put)
const mockedPatch = vi.mocked(apiClient.patch)
const mockedDelete = vi.mocked(apiClient.delete)

const sampleApiMedia: NewsletterApiMedia = {
  id: 17,
  display_name: 'Newsletter PDF',
  file_name: 'july-newsletter.pdf',
  gcp_object_key: 'news-letters/documents/july-newsletter.pdf',
  file_url: 'gs://bucket/news-letters/documents/july-newsletter.pdf',
  mime_type: 'application/pdf',
  file_size: 1024,
  media_role: 'attachment',
  sort_order: 0,
  created_at: '2026-01-01T00:00:00.000Z',
  updated_at: '2026-01-02T00:00:00.000Z',
}

const sampleApiEntry: NewsletterApiEntry = {
  id: 123,
  title: 'July Update',
  category: 'csaa',
  send_date: '2026-07-15',
  content_html: '<p>Hello members</p>',
  status: 'published',
  visibility: 'public',
  publish_at: null,
  created_at: '2026-01-01T00:00:00.000Z',
  updated_at: '2026-01-02T00:00:00.000Z',
}

const sampleDetailEntry: NewsletterApiDetailEntry = {
  ...sampleApiEntry,
  media: [sampleApiMedia],
}

function parseMultipartPayload(formData: FormData) {
  return JSON.parse(String(formData.get('payload')))
}

describe('newslettersApi converters', () => {
  it('maps newsletter entries from API shape to local shape', () => {
    const local = newsletterApiEntryToLocal(sampleApiEntry)

    expect(local).toMatchObject({
      id: '123',
      title: 'July Update',
      category: 'csaa',
      sendDate: '2026-07-15',
      contentHtml: '<p>Hello members</p>',
      status: 'published',
      visibility: 'public',
      publishAt: null,
      media: [],
    })
  })

  it('normalizes datetime send dates and unknown categories for local forms', () => {
    const local = newsletterApiEntryToLocal({
      ...sampleApiEntry,
      category: 'legacy-category' as NewsletterApiEntry['category'],
      send_date: '2026-07-15T09:30:00Z',
    })

    expect(local.category).toBe('')
    expect(local.sendDate).toBe('2026-07-15')
  })

  it('maps newsletter media to the local media shape', () => {
    const local = newsletterApiMediaToLocal(sampleApiMedia)

    expect(local).toEqual({
      id: '17',
      fileName: 'july-newsletter.pdf',
      fileUrl: 'gs://bucket/news-letters/documents/july-newsletter.pdf',
      mimeType: 'application/pdf',
      fileSize: 1024,
    })
  })

  it('uses display names or generated names when file names are missing', () => {
    const fromDisplayName = newsletterApiMediaToLocal({
      ...sampleApiMedia,
      file_name: '',
    })
    const fallback = newsletterApiMediaToLocal({
      ...sampleApiMedia,
      id: 55,
      file_name: '',
      display_name: '',
    })

    expect(fromDisplayName.fileName).toBe('Newsletter PDF')
    expect(fallback.fileName).toBe('newsletter-media-55')
  })

  it('maps detail entries including media', () => {
    const local = newsletterApiDetailToLocal(sampleDetailEntry)

    expect(local.media).toHaveLength(1)
    expect(local.media[0]).toMatchObject({
      id: '17',
      fileName: 'july-newsletter.pdf',
    })
  })
})

describe('newslettersApi CRUD', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('fetches newsletter entries and forwards list params', async () => {
    mockedGet.mockResolvedValue({
      data: { items: [sampleApiEntry], total: 1, page: 2, page_size: 5, total_pages: 1 },
    })

    const result = await fetchNewsletterEntries({ page: 2, page_size: 5, search: 'July' })

    expect(result.total).toBe(1)
    const [, config] = mockedGet.mock.calls[0]
    expect((config as { params?: Record<string, unknown> })?.params).toMatchObject({
      page: 2,
      page_size: 5,
      search: 'July',
    })
  })

  it('fetches a detail entry and maps it to the local shape', async () => {
    mockedGet.mockResolvedValue({ data: sampleDetailEntry })

    const result = await fetchNewsletterEntry('123')

    expect(result).toMatchObject({
      id: '123',
      title: 'July Update',
      media: [{ id: '17', fileName: 'july-newsletter.pdf' }],
    })
  })

  it('creates newsletter entries with JSON payloads', async () => {
    mockedPost.mockResolvedValue({
      data: {
        message: 'created',
        entry: {
          id: 123,
          title: 'July Update',
          category: 'csaa',
          send_date: '2026-07-15',
          status: 'published',
          visibility: 'public',
        },
      },
    })

    const result = await createNewsletterApiEntry({
      title: 'July Update',
      category: 'csaa',
      send_date: '2026-07-15',
      content_html: '<p>Hello members</p>',
      status: 'published',
      visibility: 'public',
      publish_at: null,
    })

    const [, body] = mockedPost.mock.calls[0]
    expect(body).toMatchObject({
      title: 'July Update',
      category: 'csaa',
      send_date: '2026-07-15',
    })
    expect(result.entry.id).toBe(123)
  })

  it('updates newsletter entries with partial JSON payloads', async () => {
    mockedPut.mockResolvedValue({
      data: {
        message: 'updated',
        entry: {
          id: 123,
          title: 'Updated July Update',
          category: 'cst',
          send_date: '2026-07-16',
          status: 'draft',
          visibility: 'private',
        },
      },
    })

    const result = await updateNewsletterApiEntry('123', {
      title: 'Updated July Update',
      category: 'cst',
      send_date: '2026-07-16',
    })

    const [url, body] = mockedPut.mock.calls[0]
    expect(String(url)).toContain('/123')
    expect(body).toEqual({
      title: 'Updated July Update',
      category: 'cst',
      send_date: '2026-07-16',
    })
    expect(result.entry.title).toBe('Updated July Update')
  })

  it('deletes a newsletter entry by id', async () => {
    mockedDelete.mockResolvedValue({})

    await deleteNewsletterApiEntry('123')

    const [url] = mockedDelete.mock.calls[0]
    expect(String(url)).toContain('/123')
  })
})

describe('newslettersApi media endpoints', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('uploads newsletter media with indexed multipart file fields', async () => {
    mockedPost.mockResolvedValue({ data: { message: 'uploaded', uploadedCount: 2 } })

    const files = [
      new File(['pdf-one'], 'july-newsletter.pdf', { type: 'application/pdf' }),
      new File(['pdf-two'], 'august-newsletter.pdf', { type: 'application/pdf' }),
    ]

    const result = await addNewsletterMedia('123', files, [
      { display_name: 'July PDF' },
      { display_name: 'August PDF', file_name: 'august-custom.pdf' },
    ])

    const [url, body] = mockedPost.mock.calls[0]
    expect(String(url)).toContain('/123/media')
    expect(body).toBeInstanceOf(FormData)
    expect(parseMultipartPayload(body as FormData)).toEqual({
      media: [
        { display_name: 'July PDF', file_name: 'july-newsletter.pdf' },
        { display_name: 'August PDF', file_name: 'august-custom.pdf' },
      ],
    })
    expect(((body as FormData).get('media[0].file') as File).name).toBe(
      'july-newsletter.pdf',
    )
    expect(((body as FormData).get('media[1].file') as File).name).toBe(
      'august-newsletter.pdf',
    )
    expect(result.uploadedCount).toBe(2)
  })

  it('updates media metadata and maps the response', async () => {
    mockedPatch.mockResolvedValue({ data: { message: 'updated', media: sampleApiMedia } })

    const result = await updateNewsletterMedia('123', '17', {
      display_name: 'July PDF',
      file_name: 'july-newsletter.pdf',
    })

    const [url, body] = mockedPatch.mock.calls[0]
    expect(String(url)).toContain('/123/media/17')
    expect(body).toEqual({
      display_name: 'July PDF',
      file_name: 'july-newsletter.pdf',
    })
    expect(result).toMatchObject({ id: '17', fileName: 'july-newsletter.pdf' })
  })

  it('downloads newsletter media as a blob', async () => {
    const blob = new Blob(['newsletter'])
    mockedGet.mockResolvedValue({ data: blob })

    const result = await getNewsletterMediaContent('123', '17')

    const [url, config] = mockedGet.mock.calls[0]
    expect(String(url)).toContain('/123/media/17/content')
    expect((config as { responseType?: string })?.responseType).toBe('blob')
    expect(result).toBe(blob)
  })

  it('reorders newsletter media with media_ids payload', async () => {
    mockedPut.mockResolvedValue({ data: { message: 'reordered', updatedCount: 2 } })

    const result = await reorderNewsletterMedia('123', [17, 18])

    const [url, body] = mockedPut.mock.calls[0]
    expect(String(url)).toContain('/123/media/order')
    expect(body).toEqual({ media_ids: [17, 18] })
    expect(result.updatedCount).toBe(2)
  })

  it('deletes newsletter media with a request body', async () => {
    mockedDelete.mockResolvedValue({ data: { message: 'deleted', deletedCount: 2 } })

    const result = await deleteNewsletterMedia('123', [17, 18])

    const [url, config] = mockedDelete.mock.calls[0]
    expect(String(url)).toContain('/123/media')
    expect((config as { data?: unknown })?.data).toEqual({ media_ids: [17, 18] })
    expect(result.deletedCount).toBe(2)
  })
})
