import { API_ROUTES } from '../constants/api'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type PageStatus = 'draft' | 'published'
export type PageStatusFilter = '' | PageStatus
export type PageSortBy =
  | 'page_title'
  | 'url_slug'
  | 'status'
  | 'last_modified'
  | 'created_at'
  | 'updated_at'
export type PageSortOrder = 'asc' | 'desc'

<<<<<<< HEAD
<<<<<<< HEAD
export type PageListItem = {
=======
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
export type PageParentReference = {
  id: number
  page_title: string
  url_slug: string
}

type PageParentRelation = {
  parent_page_id?: number | null
  parent_id?: number | null
  parent_page?: PageParentReference | null
  parent?: PageParentReference | null
}

export type PageListItem = PageParentRelation & {
<<<<<<< HEAD
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  id: number
  page_title: string
  url_slug: string
  status: PageStatus
  last_modified: string
  modified_by?: number | null
  modified_by_name: string
  created_at: string
  updated_at: string
}

export type PageListPageMeta = {
  page: number
  page_size: number
  total_items: number
  total_pages: number
  has_next: boolean
  has_prev: boolean
}

export type PageListFilters = {
  page: number
  pageSize: number
  searchTerm: string
  status: PageStatusFilter
  sortBy: PageSortBy
  sortOrder: PageSortOrder
}

export type PageListResponse = {
  items: PageListItem[]
  pagination: PageListPageMeta
  applied_filters: {
    page: number
    page_size: number
    search_term: string
    status: PageStatusFilter
    sort_by: PageSortBy
    sort_order: PageSortOrder
  }
}

type RawPageListResponse = Omit<PageListResponse, 'items'> & {
  items: PageListItem[] | null
}

<<<<<<< HEAD
<<<<<<< HEAD
export type PageDetailResponse = {
=======
export type PageDetailResponse = PageParentRelation & {
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
export type PageDetailResponse = PageParentRelation & {
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  id: number
  page_title: string
  url_slug: string
  status: PageStatus
  hero_image_enabled: boolean
  hero_image_url: string
  hero_image_object_key: string
  hero_image_fetch_url: string
  seo_page_title: string
  seo_page_description: string
  created_by?: number | null
  created_by_name: string
  modified_by?: number | null
  modified_by_name: string
  last_modified: string
  created_at: string
  updated_at: string
}

export type PageUploadInput = {
  file_name?: string
  mime_type?: string
  file_url?: string
  object_key?: string
  gcp_object_key?: string
  storage_uri?: string
}

export type SavePagePayload = {
  page_title: string
  url_slug: string
<<<<<<< HEAD
<<<<<<< HEAD
=======
  parent_page_id: number | null
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
  parent_page_id: number | null
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  status: PageStatus
  hero_image_enabled: boolean
  hero_image?: PageUploadInput
  remove_hero_image: boolean
  seo_page_title: string
  seo_page_description: string
}

export type SavePageRequest = SavePagePayload & {
  heroImageFile?: File
}

<<<<<<< HEAD
<<<<<<< HEAD
=======
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
export type PageParentOption = {
  id: number
  page_title: string
  url_slug: string
  parent_page_id: number | null
}

<<<<<<< HEAD
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
export type PageMutationResponse = {
  message: string
  page: {
    id: number
    page_title: string
    url_slug: string
    status: PageStatus
  }
}

function buildListQuery(filters: PageListFilters) {
  const params = new URLSearchParams()
<<<<<<< HEAD
<<<<<<< HEAD
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  params.set('sort_by', filters.sortBy)
  params.set('sort_order', filters.sortOrder)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }
  if (filters.status) {
    params.set('status', filters.status)
  }

  return params
}

<<<<<<< HEAD
<<<<<<< HEAD
export const pagesApi = {
  async listPages(filters: PageListFilters) {
    const response = await apiClient.get<RawPageListResponse>(API_ROUTES.pages, {
      params: buildListQuery(filters),
    })
    return {
      ...response.data,
      items: Array.isArray(response.data.items) ? response.data.items : [],
=======
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
async function fetchAllPageListItems(filters: PageListFilters) {
  const response = await apiClient.get<RawPageListResponse>(API_ROUTES.pages, {
    params: buildListQuery(filters),
  })

  return Array.isArray(response.data.items) ? response.data.items : []
}

export function resolvePageParentId(page: PageParentRelation) {
  if (typeof page.parent_page_id === 'number') {
    return page.parent_page_id
  }

  if (typeof page.parent_id === 'number') {
    return page.parent_id
  }

  if (typeof page.parent_page?.id === 'number') {
    return page.parent_page.id
  }

  if (typeof page.parent?.id === 'number') {
    return page.parent.id
  }

  return null
}

export function resolvePageParentSlug(page: PageParentRelation) {
  return page.parent_page?.url_slug ?? page.parent?.url_slug ?? ''
}

export const pagesApi = {
  async listPages(filters: PageListFilters) {
    const items = await fetchAllPageListItems(filters)
    const totalItems = items.length

    return {
      items,
      pagination: {
        page: 1,
        page_size: totalItems,
        total_items: totalItems,
        total_pages: 1,
        has_next: false,
        has_prev: false,
      },
      applied_filters: {
        page: 1,
        page_size: totalItems,
        search_term: filters.searchTerm.trim(),
        status: filters.status,
        sort_by: filters.sortBy,
        sort_order: filters.sortOrder,
      },
<<<<<<< HEAD
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
    } satisfies PageListResponse
  },

  async getPage(id: number) {
    const response = await apiClient.get<PageDetailResponse>(API_ROUTES.pageById(id))
    return response.data
  },

<<<<<<< HEAD
<<<<<<< HEAD
=======
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  async listPageParentOptions() {
    const allItems = await fetchAllPageListItems({
      page: 1,
      pageSize: 100,
      searchTerm: '',
      status: '',
      sortBy: 'page_title',
      sortOrder: 'asc',
    })

    const missingParentDataIds = allItems
      .filter(
        (item) =>
          item.parent_page_id === undefined &&
          item.parent_id === undefined &&
          item.parent_page === undefined &&
          item.parent === undefined,
      )
      .map((item) => item.id)

    const detailedPages = new Map<number, PageDetailResponse>()

    if (missingParentDataIds.length > 0) {
      const details = await Promise.all(
        missingParentDataIds.map((id) => pagesApi.getPage(id)),
      )

      for (const detail of details) {
        detailedPages.set(detail.id, detail)
      }
    }

    return allItems.map((item) => {
      const source = detailedPages.get(item.id) ?? item

      return {
        id: item.id,
        page_title: item.page_title,
        url_slug: item.url_slug,
        parent_page_id: resolvePageParentId(source),
      } satisfies PageParentOption
    })
  },

<<<<<<< HEAD
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  async fetchPageHeroImageContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createPage(request: SavePageRequest) {
    const { heroImageFile, ...payload } = request
    const body = heroImageFile
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'hero_image_file',
            file: heroImageFile,
            fileName: heroImageFile.name,
          },
        ])
      : payload
    const response = await apiClient.post<PageMutationResponse>(API_ROUTES.pages, body)
    return response.data
  },

  async updatePage(id: number, request: SavePageRequest) {
    const { heroImageFile, ...payload } = request
    const body = heroImageFile
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'hero_image_file',
            file: heroImageFile,
            fileName: heroImageFile.name,
          },
        ])
      : payload
    const response = await apiClient.put<PageMutationResponse>(API_ROUTES.pageById(id), body)
    return response.data
  },

  async deletePage(id: number) {
    await apiClient.delete(API_ROUTES.pageById(id))
  },
}
