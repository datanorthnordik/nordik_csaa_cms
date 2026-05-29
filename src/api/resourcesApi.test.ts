import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import { resourcesApi } from './resourcesApi'

vi.mock('./apiClient', () => ({
  apiClient: {
    delete: vi.fn(),
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
  },
}))

const mockedPost = vi.mocked(apiClient.post)
const mockedPut = vi.mocked(apiClient.put)

describe('resourcesApi upload validation', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('rejects unsupported files before creating a resource upload request', async () => {
    await expect(
      resourcesApi.createResource(
        {
          name: 'Test resource',
          description: 'Testing upload validation',
          category: 'media',
          visibility: 'public',
        },
        new File(['video'], 'clip.mp4', { type: 'video/mp4' }),
      ),
    ).rejects.toThrow(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )

    expect(mockedPost).not.toHaveBeenCalled()
  })

  it('rejects unsupported files before updating a resource upload request', async () => {
    await expect(
      resourcesApi.updateResource(
        '17',
        {
          name: 'Test resource',
          description: 'Testing upload validation',
          category: 'media',
          visibility: 'public',
        },
        new File(['audio'], 'voice.mp3', { type: 'audio/mpeg' }),
      ),
    ).rejects.toThrow(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )

    expect(mockedPut).not.toHaveBeenCalled()
  })
})
