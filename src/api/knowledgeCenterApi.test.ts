import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from './apiClient'
import { knowledgeCenterApi } from './knowledgeCenterApi'

vi.mock('./apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)
const mockedPost = vi.mocked(apiClient.post)

describe('knowledgeCenterApi', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('builds the list query for submission requests', async () => {
    mockedGet.mockResolvedValue({
      data: {
        items: [],
        pagination: {
          page: 2,
          page_size: 10,
          total_items: 0,
          total_pages: 0,
          has_next: false,
          has_prev: true,
        },
        summary: {
          open_count: 3,
          completed_count: 1,
        },
        applied_filters: {
          page: 2,
          page_size: 10,
          search_term: 'alice',
          status: 'open',
        },
      },
    })

    await knowledgeCenterApi.listSubmissions({
      page: 2,
      pageSize: 10,
      searchTerm: 'alice',
      status: 'open',
    })

    const [, config] = mockedGet.mock.calls[0] ?? []
    const params = config?.params as URLSearchParams
    expect(params.get('page')).toBe('2')
    expect(params.get('page_size')).toBe('10')
    expect(params.get('status')).toBe('open')
    expect(params.get('search')).toBe('alice')
  })

  it('posts completion notes when marking a request complete', async () => {
    mockedPost.mockResolvedValue({
      data: {
        submission: {
          id: 17,
          submitter_name: 'Alice',
          submitter_email: 'alice@example.com',
          submitter_phone: '',
          submission_type: 'post',
          message: 'Story',
          status: 'completed',
          completion_notes: 'Published on the site.',
          completed_by: {
            id: 8,
            name: 'Jane Doe',
            email: 'jane@example.com',
          },
          completed_at: '2026-06-18T14:00:00Z',
          created_at: '2026-06-18T12:00:00Z',
          updated_at: '2026-06-18T14:00:00Z',
        },
      },
    })

    await knowledgeCenterApi.completeSubmission(
      17,
      'Published on the site.',
    )

    expect(mockedPost).toHaveBeenCalledWith(
      '/api/knowledge-center/submissions/17/complete',
      {
        completion_notes: 'Published on the site.',
      },
      {
        skipErrorToast: true,
      },
    )
  })
})
