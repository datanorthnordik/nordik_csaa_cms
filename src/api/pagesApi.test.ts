import { beforeEach, describe, expect, it, vi } from 'vitest'
import { pagesApi, type PageListItem } from './pagesApi'

const { getMock } = vi.hoisted(() => ({
  getMock: vi.fn(),
}))

vi.mock('./apiClient', () => ({
  apiClient: {
    get: getMock,
  },
}))

function createPageListItem(id: number): PageListItem {
  return {
    id,
    page_title: `Page ${id}`,
    url_slug: `/page-${id}`,
    parent_page_title: '',
    parent_page_url_slug: '',
    status: 'draft',
    last_modified: '2026-05-12T00:00:00Z',
    modified_by: null,
    modified_by_name: 'Admin',
    created_at: '2026-05-12T00:00:00Z',
    updated_at: '2026-05-12T00:00:00Z',
  }
}

describe('pagesApi.listPages', () => {
  beforeEach(() => {
    getMock.mockReset()
  })

  it('omits page and page_size from the request and returns all filtered pages', async () => {
    getMock.mockResolvedValue({
      data: {
        items: Array.from({ length: 12 }, (_, index) => createPageListItem(index + 1)),
        pagination: {
          page: 1,
          page_size: 12,
          total_items: 12,
          total_pages: 1,
          has_next: false,
          has_prev: false,
        },
        applied_filters: {
          page: 1,
          page_size: 12,
          search_term: '',
          status: '',
          sort_by: 'last_modified',
          sort_order: 'desc',
        },
      },
    })

    const response = await pagesApi.listPages({
      page: 2,
      pageSize: 5,
      searchTerm: '',
      status: 'draft',
      sortBy: 'last_modified',
      sortOrder: 'desc',
    })

    expect(getMock).toHaveBeenCalledTimes(1)
    expect(getMock.mock.calls[0]?.[1]?.params.toString()).toBe(
      'sort_by=last_modified&sort_order=desc&status=draft',
    )
    expect(response.items.map((item) => item.id)).toEqual([
      1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12,
    ])
    expect(response.pagination).toEqual({
      page: 1,
      page_size: 12,
      total_items: 12,
      total_pages: 1,
      has_next: false,
      has_prev: false,
    })
    expect(response.applied_filters).toEqual({
      page: 1,
      page_size: 12,
      search_term: '',
      status: 'draft',
      sort_by: 'last_modified',
      sort_order: 'desc',
    })
  })
})

describe('pagesApi.listPageParentOptions', () => {
  beforeEach(() => {
    getMock.mockReset()
  })

  it('returns parent ids from the page list without fetching every page detail', async () => {
    getMock.mockResolvedValue({
      data: {
        items: [
          {
            ...createPageListItem(1),
            page_title: 'Home',
            url_slug: '/home',
            parent_id: null,
          },
          {
            ...createPageListItem(4),
            page_title: 'Contact Us',
            url_slug: '/home/contact-us',
            parent_id: 1,
          },
          {
            ...createPageListItem(6),
            page_title: 'Test1',
            url_slug: '/home/contact-us/test1',
            parent_id: 4,
          },
        ],
        pagination: {
          page: 1,
          page_size: 3,
          total_items: 3,
          total_pages: 1,
          has_next: false,
          has_prev: false,
        },
        applied_filters: {
          page: 0,
          page_size: 0,
          search_term: '',
          status: '',
          sort_by: 'page_title',
          sort_order: 'asc',
        },
      },
    })

    const response = await pagesApi.listPageParentOptions()

    expect(getMock).toHaveBeenCalledTimes(1)
    expect(response).toEqual([
      {
        id: 1,
        page_title: 'Home',
        url_slug: '/home',
        parent_id: null,
      },
      {
        id: 4,
        page_title: 'Contact Us',
        url_slug: '/home/contact-us',
        parent_id: 1,
      },
      {
        id: 6,
        page_title: 'Test1',
        url_slug: '/home/contact-us/test1',
        parent_id: 4,
      },
    ])
  })
})
