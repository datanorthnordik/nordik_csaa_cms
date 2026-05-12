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

export type PageListItem = {
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

export type PageDetailResponse = {
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
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
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

export const pagesApi = {
  async listPages(filters: PageListFilters) {
    const response = await apiClient.get<RawPageListResponse>(API_ROUTES.pages, {
      params: buildListQuery(filters),
    })
    return {
      ...response.data,
      items: Array.isArray(response.data.items) ? response.data.items : [],
    } satisfies PageListResponse
  },

  async getPage(id: number) {
    const response = await apiClient.get<PageDetailResponse>(API_ROUTES.pageById(id))
    return response.data
  },

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
