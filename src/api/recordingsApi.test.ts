import { beforeEach, describe, expect, it, vi } from 'vitest'
import { API_BASE_URL, API_ROUTES } from '../constants/api'

const { deleteMock, getMock, patchMock, postMock, putMock } = vi.hoisted(() => ({
  deleteMock: vi.fn(),
  getMock: vi.fn(),
  patchMock: vi.fn(),
  postMock: vi.fn(),
  putMock: vi.fn(),
}))

vi.mock('./apiClient', () => ({
  apiClient: {
    delete: deleteMock,
    get: getMock,
    patch: patchMock,
    post: postMock,
    put: putMock,
  },
}))

import { recordingsApi } from './recordingsApi'

describe('recordingsApi', () => {
  beforeEach(() => {
    deleteMock.mockReset()
    getMock.mockReset()
    patchMock.mockReset()
    postMock.mockReset()
    putMock.mockReset()
  })

  it('maps recording collection summaries and preserves list order', async () => {
    getMock.mockResolvedValue({
      data: {
        items: [
          {
            id: 2,
            name: 'Board Meetings',
            item_count: 4,
            created_at: '2026-06-15T00:00:00Z',
            updated_at: '2026-06-15T01:00:00Z',
          },
          {
            id: 1,
            name: 'Annual General Meeting',
            item_count: 0,
            created_at: '2026-06-14T00:00:00Z',
            updated_at: '2026-06-14T01:00:00Z',
          },
        ],
      },
    })

    const response = await recordingsApi.listRecordingCollections()

    expect(getMock).toHaveBeenCalledWith(API_ROUTES.recordings)
    expect(response).toEqual([
      {
        id: 2,
        name: 'Board Meetings',
        itemCount: 4,
        createdAt: '2026-06-15T00:00:00Z',
        updatedAt: '2026-06-15T01:00:00Z',
      },
      {
        id: 1,
        name: 'Annual General Meeting',
        itemCount: 0,
        createdAt: '2026-06-14T00:00:00Z',
        updatedAt: '2026-06-14T01:00:00Z',
      },
    ])
  })

  it('defaults the collection list to an empty array when the payload omits items', async () => {
    getMock.mockResolvedValue({ data: {} })

    await expect(recordingsApi.listRecordingCollections()).resolves.toEqual([])
  })

  it('maps a recording collection detail with items and defaults missing fields', async () => {
    getMock.mockResolvedValue({
      data: {
        id: 7,
        name: 'Town Halls',
        item_count: 1,
        items: [
          {
            id: 14,
            recording_collection_id: 7,
            title: 'Spring Town Hall',
            description: null,
            recording_url: null,
            sort_order: 0,
            created_at: '2026-06-15T00:00:00Z',
            updated_at: '2026-06-15T01:00:00Z',
          },
        ],
        created_at: '2026-06-15T00:00:00Z',
        updated_at: '2026-06-15T01:00:00Z',
      },
    })

    const response = await recordingsApi.getRecordingCollection(7)

    expect(getMock).toHaveBeenCalledWith(API_ROUTES.recordingById(7))
    expect(response).toEqual({
      id: 7,
      name: 'Town Halls',
      itemCount: 1,
      items: [
        {
          id: 14,
          recordingCollectionId: 7,
          title: 'Spring Town Hall',
          description: '',
          recordingUrl: undefined,
          sortOrder: 0,
          createdAt: '2026-06-15T00:00:00Z',
          updatedAt: '2026-06-15T01:00:00Z',
        },
      ],
      createdAt: '2026-06-15T00:00:00Z',
      updatedAt: '2026-06-15T01:00:00Z',
    })
  })

  it('creates a recording collection with item audio as multipart data', async () => {
    postMock.mockResolvedValue({
      data: {
        message: 'Recording collection created successfully',
        recording: { id: 9, name: 'Weekly Updates' },
      },
    })

    const recordingFile = new File(['audio'], 'week-1.mp3', { type: 'audio/mpeg' })
    const response = await recordingsApi.createRecordingCollection({
      name: 'Weekly Updates',
      items: [{ title: 'Week 1', description: 'Kickoff' }],
      recordingFiles: [recordingFile],
    })

    expect(postMock).toHaveBeenCalledWith(API_ROUTES.recordings, expect.any(FormData))
    const formData = postMock.mock.calls[0][1] as FormData
    expect(JSON.parse(String(formData.get('payload')))).toEqual({
      name: 'Weekly Updates',
      items: [{ title: 'Week 1', description: 'Kickoff' }],
    })
    const uploadedFile = formData.get('items[0].recording_file') as File
    expect(uploadedFile.name).toBe('week-1.mp3')
    expect(uploadedFile.type).toBe('audio/mpeg')
    expect(response.recording).toEqual({ id: 9, name: 'Weekly Updates' })
  })

  it('updates the recording collection name via PUT', async () => {
    putMock.mockResolvedValue({
      data: {
        message: 'Recording collection updated successfully',
        recording: { id: 9, name: 'Renamed Collection' },
      },
    })

    await recordingsApi.updateRecordingCollection(9, { name: 'Renamed Collection' })

    expect(putMock).toHaveBeenCalledWith(API_ROUTES.recordingById(9), {
      name: 'Renamed Collection',
    })
  })

  it('deletes a recording collection via DELETE', async () => {
    deleteMock.mockResolvedValue({ data: { message: 'Recording collection deleted' } })

    await recordingsApi.deleteRecordingCollection(9)

    expect(deleteMock).toHaveBeenCalledWith(API_ROUTES.recordingById(9))
  })

  it('adds recording items to a collection', async () => {
    postMock.mockResolvedValue({
      data: { message: 'Recording items added', uploadedCount: 2 },
    })

    await recordingsApi.addRecordingItems(9, {
      items: [
        { title: 'Item One', description: 'First' },
        { title: 'Item Two' },
      ],
    })

    expect(postMock).toHaveBeenCalledWith(API_ROUTES.recordingItemsById(9), {
      items: [
        { title: 'Item One', description: 'First' },
        { title: 'Item Two' },
      ],
    })
  })

  it('patches a recording item and maps the returned item', async () => {
    patchMock.mockResolvedValue({
      data: {
        message: 'Recording item updated successfully',
        item: {
          id: 14,
          recording_collection_id: 9,
          title: 'Updated Item',
          description: 'Updated description',
          recording_url: '/api/recordings/9/items/14/content',
          sort_order: 1,
          created_at: '2026-06-15T00:00:00Z',
          updated_at: '2026-06-15T01:00:00Z',
        },
      },
    })

    const replacementFile = new File(['replacement'], 'updated.m4a', { type: 'audio/mp4' })
    const response = await recordingsApi.updateRecordingItem(9, 14, {
      title: 'Updated Item',
      description: 'Updated description',
      recordingFile: replacementFile,
    })

    expect(patchMock).toHaveBeenCalledWith(API_ROUTES.recordingItemById(9, 14), expect.any(FormData))
    const formData = patchMock.mock.calls[0][1] as FormData
    expect(JSON.parse(String(formData.get('payload')))).toEqual({
      title: 'Updated Item',
      description: 'Updated description',
    })
    const uploadedFile = formData.get('recording_file') as File
    expect(uploadedFile.name).toBe('updated.m4a')
    expect(uploadedFile.type).toBe('audio/mp4')
    expect(response.item).toEqual({
      id: 14,
      recordingCollectionId: 9,
      title: 'Updated Item',
      description: 'Updated description',
      recordingUrl: `${API_BASE_URL}/api/recordings/9/items/14/content`,
      sortOrder: 1,
      createdAt: '2026-06-15T00:00:00Z',
      updatedAt: '2026-06-15T01:00:00Z',
    })
  })

  it('deletes a recording item via DELETE', async () => {
    deleteMock.mockResolvedValue({
      data: { message: 'Recording item deleted', deletedCount: 1 },
    })

    await recordingsApi.deleteRecordingItem(9, 14)

    expect(deleteMock).toHaveBeenCalledWith(API_ROUTES.recordingItemById(9, 14))
  })
})
