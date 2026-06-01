import { beforeEach, describe, expect, it, vi } from 'vitest'
import { RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES } from '../lib/resourceUpload'
import { apiClient } from './apiClient'
import {
  addPressMedia,
  createPressApiEntry,
  deletePressApiEntry,
  deletePressMedia,
  fetchPressCoverImageContent,
  fetchPressEntries,
  fetchPressEntry,
  getPressMediaContent,
  localEntryToPressApiInput,
  pressApiDetailToLocal,
  pressApiEntryToLocal,
  pressApiMediaToLocal,
  reorderPressMedia,
  updatePressApiEntry,
  updatePressMedia,
  type PressApiDetailEntry,
  type PressApiEntry,
  type PressApiMedia,
} from './pressApi'

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

const sampleApiMedia: PressApiMedia = {
  id: 17,
  display_name: 'Agenda',
  file_name: 'agenda.pdf',
  gcp_object_key: 'press/agenda.pdf',
  file_url: 'gs://bucket/press/agenda.pdf',
  mime_type: 'application/pdf',
  file_size: 1024,
  media_role: 'attachment',
  sort_order: 0,
  created_at: '2026-01-01T00:00:00.000Z',
  updated_at: '2026-01-02T00:00:00.000Z',
}

const sampleApiEntry: PressApiEntry = {
  id: 123,
  title: 'Spring Release',
  release_date: '2026-03-20',
  category_id: 7,
  source_url: 'https://example.com/news',
  content_html: '<p>Body</p>',
  status: 'published',
  visibility: 'public',
  publish_at: null,
  cover_image_url: 'https://cdn.example.com/cover.png',
  created_at: '2026-01-01T00:00:00.000Z',
  updated_at: '2026-01-02T00:00:00.000Z',
}

const sampleDetailEntry: PressApiDetailEntry = {
  ...sampleApiEntry,
  media: [sampleApiMedia],
}

function parseMultipartPayload(formData: FormData) {
  return JSON.parse(String(formData.get('payload')))
}

describe('pressApi converters', () => {
  it('maps press entry fields from snake_case to local shape', () => {
    const local = pressApiEntryToLocal(sampleApiEntry)

    expect(local).toMatchObject({
      id: '123',
      title: 'Spring Release',
      releaseDate: '2026-03-20',
      categoryId: '7',
      sourceUrl: 'https://example.com/news',
      contentHtml: '<p>Body</p>',
      status: 'published',
      visibility: 'public',
      publishAt: null,
      coverImageUrl: 'https://cdn.example.com/cover.png',
      media: [],
    })
  })

  it('normalizes datetime release values for date inputs', () => {
    const local = pressApiEntryToLocal({
      ...sampleApiEntry,
      release_date: '2026-03-20T00:00:00Z',
    })

    expect(local.releaseDate).toBe('2026-03-20')
  })

  it('maps press media fields to the local media shape', () => {
    const local = pressApiMediaToLocal(sampleApiMedia)

    expect(local).toEqual({
      id: '17',
      fileName: 'agenda.pdf',
      fileUrl: 'gs://bucket/press/agenda.pdf',
      mimeType: 'application/pdf',
      fileSize: 1024,
    })
  })

  it('maps detail entries including media', () => {
    const local = pressApiDetailToLocal(sampleDetailEntry)

    expect(local.media).toHaveLength(1)
    expect(local.media[0]).toMatchObject({
      id: '17',
      fileName: 'agenda.pdf',
    })
  })

  it('maps local entries back to API save input', () => {
    const local = pressApiDetailToLocal(sampleDetailEntry)
    const input = localEntryToPressApiInput(local)

    expect(input).toEqual({
      title: 'Spring Release',
      release_date: '2026-03-20',
      category_id: 7,
      source_url: 'https://example.com/news',
      content_html: '<p>Body</p>',
      status: 'published',
      visibility: 'public',
      publish_at: null,
    })
  })
})

describe('pressApi CRUD', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('fetches press entries and returns the list response', async () => {
    mockedGet.mockResolvedValue({
      data: { items: [sampleApiEntry], total: 1, page: 1, page_size: 10, total_pages: 1 },
    })

    const result = await fetchPressEntries({ page: 2, page_size: 5 })

    expect(result.total).toBe(1)
    const [, config] = mockedGet.mock.calls[0]
    expect((config as any)?.params?.page).toBe(2)
    expect((config as any)?.params?.page_size).toBe(5)
  })

  it('fetches a detail entry and maps it to the local shape', async () => {
    mockedGet.mockResolvedValue({ data: sampleDetailEntry })

    const result = await fetchPressEntry('123')

    expect(result).toMatchObject({
      id: '123',
      title: 'Spring Release',
      media: [{ id: '17', fileName: 'agenda.pdf' }],
    })
  })

  it('posts JSON entry data when no cover image file is provided', async () => {
    mockedPost.mockResolvedValue({
      data: { message: 'created', entry: { id: 123, title: 'Spring Release', release_date: '2026-03-20', status: 'published', visibility: 'public' } },
    })

    const result = await createPressApiEntry({
      title: 'Spring Release',
      release_date: '2026-03-20',
      category_id: 7,
      source_url: 'https://example.com/news',
      content_html: '<p>Body</p>',
      status: 'published',
      visibility: 'public',
      publish_at: null,
    })

    const [, body] = mockedPost.mock.calls[0]
    expect(body).toMatchObject({ title: 'Spring Release', category_id: 7 })
    expect(result.entry.id).toBe(123)
  })

  it('posts multipart entry data when a cover image file is provided', async () => {
    mockedPost.mockResolvedValue({
      data: { message: 'created', entry: { id: 123, title: 'Spring Release', release_date: '2026-03-20', status: 'published', visibility: 'public' } },
    })

    const coverImage = new File(['cover'], 'cover.png', { type: 'image/png' })
    await createPressApiEntry(
      {
        title: 'Spring Release',
        release_date: '2026-03-20',
        category_id: null,
        source_url: '',
        content_html: '<p>Body</p>',
        status: 'draft',
        visibility: 'private',
        publish_at: null,
      },
      coverImage,
    )

    const [, body] = mockedPost.mock.calls[0]
    expect(body).toBeInstanceOf(FormData)
    expect(parseMultipartPayload(body as FormData)).toMatchObject({
      title: 'Spring Release',
      visibility: 'private',
    })
    expect((body as FormData).get('cover_image_file')).toBeInstanceOf(File)
    expect(((body as FormData).get('cover_image_file') as File).name).toBe('cover.png')
  })

  it('rejects unsupported cover image files before create requests are sent', async () => {
    const coverImage = new File(['cover'], 'cover.gif', { type: 'image/gif' })

    await expect(
      createPressApiEntry(
        {
          title: 'Spring Release',
          release_date: '2026-03-20',
          category_id: null,
          source_url: '',
          content_html: '<p>Body</p>',
          status: 'draft',
          visibility: 'private',
          publish_at: null,
        },
        coverImage,
      ),
    ).rejects.toThrow('Only SVG, PNG, JPG, and WEBP are supported.')

    expect(mockedPost).not.toHaveBeenCalled()
  })

  it('puts multipart update data when a new cover image is provided', async () => {
    mockedPut.mockResolvedValue({
      data: { message: 'updated', entry: { id: 123, title: 'Updated Release', release_date: '2026-03-21', status: 'draft', visibility: 'private' } },
    })

    const coverImage = new File(['cover'], 'cover.png', { type: 'image/png' })
    await updatePressApiEntry(
      '123',
      {
        title: 'Updated Release',
        release_date: '2026-03-21',
      },
      coverImage,
    )

    const [url, body] = mockedPut.mock.calls[0]
    expect(String(url)).toContain('/123')
    expect(body).toBeInstanceOf(FormData)
    expect(parseMultipartPayload(body as FormData)).toMatchObject({
      title: 'Updated Release',
      release_date: '2026-03-21',
    })
  })

  it('rejects oversized cover image files before update requests are sent', async () => {
    const coverImage = new File(['cover'], 'cover.png', { type: 'image/png' })
    Object.defineProperty(coverImage, 'size', {
      configurable: true,
      value: RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
    })

    await expect(
      updatePressApiEntry(
        '123',
        {
          title: 'Updated Release',
          release_date: '2026-03-21',
        },
        coverImage,
      ),
    ).rejects.toThrow('This file exceeds the 20MB limit.')

    expect(mockedPut).not.toHaveBeenCalled()
  })

  it('deletes a press entry by id', async () => {
    mockedDelete.mockResolvedValue({})

    await deletePressApiEntry('123')

    const [url] = mockedDelete.mock.calls[0]
    expect(String(url)).toContain('/123')
  })
})

describe('pressApi media endpoints', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('uploads media with indexed multipart file fields', async () => {
    mockedPost.mockResolvedValue({ data: { message: 'uploaded', uploadedCount: 2 } })

    const files = [
      new File(['agenda'], 'agenda.pdf', { type: 'application/pdf' }),
      new File(['minutes'], 'minutes.pdf', { type: 'application/pdf' }),
    ]

    const result = await addPressMedia('123', files, [
      { display_name: 'Agenda' },
      { display_name: 'Minutes' },
    ])

    const [url, body] = mockedPost.mock.calls[0]
    expect(String(url)).toContain('/123/media')
    expect(body).toBeInstanceOf(FormData)
    expect(parseMultipartPayload(body as FormData)).toEqual({
      media: [
        { display_name: 'Agenda', file_name: 'agenda.pdf' },
        { display_name: 'Minutes', file_name: 'minutes.pdf' },
      ],
    })
    expect((body as FormData).get('media[0].file')).toBeInstanceOf(File)
    expect(((body as FormData).get('media[0].file') as File).name).toBe('agenda.pdf')
    expect((body as FormData).get('media[1].file')).toBeInstanceOf(File)
    expect(((body as FormData).get('media[1].file') as File).name).toBe('minutes.pdf')
    expect(result.uploadedCount).toBe(2)
  })

  it('rejects unsupported media files before upload requests are sent', async () => {
    await expect(
      addPressMedia(
        '123',
        [new File(['audio'], 'voice.mp3', { type: 'audio/mpeg' })],
        [{ display_name: 'Voice memo' }],
      ),
    ).rejects.toThrow('Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.')

    expect(mockedPost).not.toHaveBeenCalled()
  })

  it('updates media metadata and maps the response', async () => {
    mockedPatch.mockResolvedValue({ data: { message: 'updated', media: sampleApiMedia } })

    const result = await updatePressMedia('123', '17', {
      display_name: 'Agenda',
      file_name: 'agenda.pdf',
    })

    const [url, body] = mockedPatch.mock.calls[0]
    expect(String(url)).toContain('/123/media/17')
    expect(body).toEqual({ display_name: 'Agenda', file_name: 'agenda.pdf' })
    expect(result).toMatchObject({ id: '17', fileName: 'agenda.pdf' })
  })

  it('downloads press media as a blob', async () => {
    const blob = new Blob(['agenda'])
    mockedGet.mockResolvedValue({ data: blob })

    const result = await getPressMediaContent('123', '17')

    const [url, config] = mockedGet.mock.calls[0]
    expect(String(url)).toContain('/123/media/17/content')
    expect((config as any)?.responseType).toBe('blob')
    expect(result).toBe(blob)
  })

  it('downloads press cover images as a blob', async () => {
    const blob = new Blob(['cover'])
    mockedGet.mockResolvedValue({ data: blob })

    const result = await fetchPressCoverImageContent('123')

    const [url, config] = mockedGet.mock.calls[0]
    expect(String(url)).toContain('/123/cover/content')
    expect((config as any)?.responseType).toBe('blob')
    expect(result).toBe(blob)
  })

  it('reorders media with media_ids payload', async () => {
    mockedPut.mockResolvedValue({ data: { message: 'reordered', updatedCount: 2 } })

    const result = await reorderPressMedia('123', [17, 18])

    const [url, body] = mockedPut.mock.calls[0]
    expect(String(url)).toContain('/123/media/order')
    expect(body).toEqual({ media_ids: [17, 18] })
    expect(result.updatedCount).toBe(2)
  })

  it('deletes media with a request body', async () => {
    mockedDelete.mockResolvedValue({ data: { message: 'deleted', deletedCount: 2 } })

    const result = await deletePressMedia('123', [17, 18])

    const [url, config] = mockedDelete.mock.calls[0]
    expect(String(url)).toContain('/123/media')
    expect((config as any)?.data).toEqual({ media_ids: [17, 18] })
    expect(result.deletedCount).toBe(2)
  })
})
