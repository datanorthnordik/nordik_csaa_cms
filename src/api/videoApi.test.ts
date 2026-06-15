import { beforeEach, describe, expect, it, vi } from 'vitest'
import { API_ROUTES } from '../constants/api'

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

import { videoApi } from './videoApi'

describe('videoApi', () => {
  beforeEach(() => {
    deleteMock.mockReset()
    getMock.mockReset()
    patchMock.mockReset()
    postMock.mockReset()
    putMock.mockReset()
  })

  it('maps video package summaries from the API', async () => {
    getMock.mockResolvedValue({
      data: {
        items: [
          {
            id: 12,
            title: 'Community Stories',
            package_type: 'collection',
            video_count: 3,
            front_image_url: '/api/videos/12/items/8/teaser/content',
            created_at: '2026-06-15T00:00:00Z',
            updated_at: '2026-06-15T01:00:00Z',
          },
        ],
      },
    })

    const response = await videoApi.listVideoPackages()

    expect(getMock).toHaveBeenCalledWith(API_ROUTES.videos)
    expect(response).toEqual([
      {
        id: 12,
        title: 'Community Stories',
        packageType: 'collection',
        videoCount: 3,
        frontImageUrl: '/api/videos/12/items/8/teaser/content',
        frontImagePath: '/api/videos/12/items/8/teaser/content',
        createdAt: '2026-06-15T00:00:00Z',
        updatedAt: '2026-06-15T01:00:00Z',
      },
    ])
  })

  it('posts single video packages as multipart payloads when a teaser file is included', async () => {
    const teaserFile = new File(['teaser'], 'teaser.png', { type: 'image/png' })
    postMock.mockResolvedValue({
      data: {
        message: 'Video package created successfully',
        video: {
          id: 5,
          title: 'Welcome Video',
          package_type: 'single',
        },
      },
    })

    await videoApi.createVideoPackage({
      title: 'Welcome Video',
      package_type: 'single',
      single_video: {
        title: 'Welcome Video',
        youtube_url: 'https://www.youtube.com/watch?v=welcome123',
        description: 'Homepage greeting',
        file_name: 'teaser.png',
        mime_type: 'image/png',
      },
      singleVideoFile: teaserFile,
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.videos)

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect(body.get('payload')).toBe(
      JSON.stringify({
        title: 'Welcome Video',
        package_type: 'single',
        single_video: {
          title: 'Welcome Video',
          youtube_url: 'https://www.youtube.com/watch?v=welcome123',
          description: 'Homepage greeting',
          file_name: 'teaser.png',
          mime_type: 'image/png',
        },
      }),
    )
    expect((body.get('single_video.teaser_image_file') as File).name).toBe('teaser.png')
  })

  it('posts collection video packages with indexed teaser image fields', async () => {
    const firstFile = new File(['first'], 'first.png', { type: 'image/png' })
    const secondFile = new File(['second'], 'second.png', { type: 'image/png' })
    postMock.mockResolvedValue({
      data: {
        message: 'Video package created successfully',
        video: {
          id: 7,
          title: 'Youth Videos',
          package_type: 'collection',
        },
      },
    })

    await videoApi.createVideoPackage({
      title: 'Youth Videos',
      package_type: 'collection',
      videos: [
        {
          title: 'Story One',
          youtube_url: 'https://youtu.be/story-one',
          description: 'First story',
          file_name: 'first.png',
          mime_type: 'image/png',
        },
        {
          title: 'Story Two',
          youtube_url: 'https://youtu.be/story-two',
          description: 'Second story',
          file_name: 'second.png',
          mime_type: 'image/png',
        },
      ],
      videoFiles: [firstFile, secondFile],
    })

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('videos[0].teaser_image_file') as File).name).toBe('first.png')
    expect((body.get('videos[1].teaser_image_file') as File).name).toBe('second.png')
  })

  it('posts add-video-item requests as multipart payloads when new teaser files are included', async () => {
    const teaserFile = new File(['addition'], 'addition.png', { type: 'image/png' })
    postMock.mockResolvedValue({
      data: {
        message: 'Video items uploaded successfully',
        uploadedCount: 1,
      },
    })

    await videoApi.addVideoItems(9, {
      videos: [
        {
          title: 'Added Video',
          youtube_url: 'https://www.youtube.com/watch?v=added123',
          description: 'A new story',
          file_name: 'addition.png',
          mime_type: 'image/png',
        },
      ],
      videoFiles: [teaserFile],
    })

    expect(postMock).toHaveBeenCalledWith(
      `${API_ROUTES.videoById(9)}/items`,
      expect.any(FormData),
    )
    const body = postMock.mock.calls[0]?.[1] as FormData
    expect((body.get('videos[0].teaser_image_file') as File).name).toBe('addition.png')
  })

  it('patches video items with teaser image multipart fields when replacing the teaser', async () => {
    const teaserFile = new File(['replacement'], 'replacement.png', { type: 'image/png' })
    patchMock.mockResolvedValue({
      data: {
        message: 'Video item updated successfully',
        item: {
          id: 14,
          video_package_id: 9,
          title: 'Updated Story',
          youtube_url: 'https://www.youtube.com/watch?v=updated123',
          description: 'Updated description',
          teaser_image_url: '/api/videos/9/items/14/teaser/content',
          storage_uri: 'gs://bucket/videos/9/items/replacement.png',
          gcp_object_key: 'videos/9/items/replacement.png',
          sort_order: 0,
          created_at: '2026-06-15T00:00:00Z',
          updated_at: '2026-06-15T01:00:00Z',
        },
      },
    })

    const response = await videoApi.updateVideoItem(9, 14, {
      title: 'Updated Story',
      youtube_url: 'https://www.youtube.com/watch?v=updated123',
      description: 'Updated description',
      file_name: 'replacement.png',
      mime_type: 'image/png',
      teaserImageFile: teaserFile,
    })

    expect(patchMock).toHaveBeenCalledWith(
      API_ROUTES.videoItemById(9, 14),
      expect.any(FormData),
    )
    const body = patchMock.mock.calls[0]?.[1] as FormData
    expect((body.get('teaser_image_file') as File).name).toBe('replacement.png')
    expect(response.item.teaserImagePath).toBe('/api/videos/9/items/14/teaser/content')
  })
})
