import { API_ROUTES } from '../constants/api'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type PageStatus = 'draft' | 'published'
export type PageType = 'page' | 'module'
export type PageStatusFilter = '' | PageStatus
export type PageSortBy =
  | 'page_title'
  | 'url_slug'
  | 'status'
  | 'last_modified'
  | 'created_at'
  | 'updated_at'
export type PageSortOrder = 'asc' | 'desc'
export type PageSectionType =
  | 'header'
  | 'typography'
  | 'gallery'
  | 'video'
  | 'document'
  | 'quote'
  | 'cta_banner'
export type PageHeaderHierarchy = 'h1_hero' | 'h2_section'
export type PageGalleryViewMode = 'grid' | 'carousel' | 'masonry' | 'focus' | 'icons'
export type PageTypographyTextAlign = 'left' | 'center' | 'right'

type JSONObject = Record<string, unknown>

type PageParentRelation = {
  parent_id?: number | null
  parent_page_title?: string
  parent_page_url_slug?: string
}

export type PageListItem = PageParentRelation & {
  id: number
  page_title: string
  url_slug: string
  page_type: PageType
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

export type PageHeaderSectionResponse = {
  main_header_text: string
  sub_header_text: string
  description?: string
  hierarchy: PageHeaderHierarchy
  text_align: PageTypographyTextAlign
  underline_enabled?: boolean
}

export type PageTypographySectionResponse = {
  html_content: string
  text_content: string
  text_align: PageTypographyTextAlign
}

export type PageGallerySectionResponse = {
  gallery_id?: number | null
  view_mode: PageGalleryViewMode
  show_title_description?: boolean
  auto_scroll_enabled?: boolean
}

export type PageVideoSectionResponse = {
  video_package_id?: number | null
}

export type PageQuoteSectionResponse = {
  quote_content: string
  attribution: string
}

export type PageSectionAssetResponse = {
  file_url: string
  fetch_url: string
  storage_uri: string
  gcp_object_key: string
}

export type PageCTABannerSectionResponse = {
  banner_heading: string
  banner_message: string
  button_text: string
  button_url: string
  open_in_new_tab: boolean
  image?: PageSectionAssetResponse | null
}

export type PageDocumentResponse = {
  id: number
  display_name: string
  description: string
  original_file_name: string
  file_name: string
  file_url: string
  fetch_url: string
  storage_uri: string
  gcp_object_key: string
  mime_type: string
  file_size: number
  sort_order: number
  created_at: string
  updated_at: string
}

export type PageDocumentsSectionResponse = {
  items: PageDocumentResponse[]
}

export type PageSectionResponse = {
  id: number
  section_name: string
  section_type: PageSectionType
  sort_order: number
  is_enabled: boolean
  settings?: JSONObject | null
  header?: PageHeaderSectionResponse | null
  typography?: PageTypographySectionResponse | null
  gallery?: PageGallerySectionResponse | null
  video?: PageVideoSectionResponse | null
  quote?: PageQuoteSectionResponse | null
  cta_banner?: PageCTABannerSectionResponse | null
  documents?: PageDocumentsSectionResponse | null
  created_at: string
  updated_at: string
}

export type PageContentDetailResponse = {
  id: number
  page_id: number
  template_key: string
  settings?: JSONObject | null
  schema_version: number
  sections: PageSectionResponse[]
  created_by?: number | null
  updated_by?: number | null
  created_at?: string
  updated_at?: string
}

export type PageDetailResponse = PageParentRelation & {
  id: number
  page_title: string
  url_slug: string
  page_type: PageType
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
  page_detail?: PageContentDetailResponse | null
}

export type PageUploadInput = {
  file_name?: string
  mime_type?: string
  file_url?: string
  object_key?: string
  gcp_object_key?: string
  storage_uri?: string
}

export type SavePageDocumentPayload = {
  id?: number
  display_name: string
  description: string
  original_file_name: string
  file_name?: string
  mime_type?: string
  file_url?: string
  object_key?: string
  gcp_object_key?: string
  storage_uri?: string
}

export type SavePageDocumentsSectionPayload = {
  items: SavePageDocumentPayload[]
}

export type SavePageHeaderSectionPayload = {
  main_header_text: string
  sub_header_text: string
  description: string
  hierarchy: PageHeaderHierarchy
  text_align: PageTypographyTextAlign
  underline_enabled: boolean
}

export type SavePageTypographySectionPayload = {
  html_content: string
  text_content: string
  text_align: PageTypographyTextAlign
}

export type SavePageGallerySectionPayload = {
  gallery_id?: number | null
  view_mode: PageGalleryViewMode
  show_title_description: boolean
  auto_scroll_enabled: boolean
}

export type SavePageVideoSectionPayload = {
  video_package_id?: number | null
}

export type SavePageQuoteSectionPayload = {
  quote_content: string
  attribution: string
}

export type SavePageCTABannerSectionPayload = {
  banner_heading: string
  banner_message: string
  button_text: string
  button_url: string
  open_in_new_tab: boolean
  image?: PageUploadInput
}

export type SavePageSectionPayload = {
  id?: number
  section_name: string
  section_type: PageSectionType
  sort_order: number
  is_enabled: boolean
  settings?: JSONObject
  header?: SavePageHeaderSectionPayload
  typography?: SavePageTypographySectionPayload
  gallery?: SavePageGallerySectionPayload
  video?: SavePageVideoSectionPayload
  quote?: SavePageQuoteSectionPayload
  cta_banner?: SavePageCTABannerSectionPayload
  documents?: SavePageDocumentsSectionPayload
}

export type SavePageDetailPayload = {
  template_key: string
  settings?: JSONObject
  sections: SavePageSectionPayload[]
}

export type SavePagePayload = {
  page_title: string
  url_slug: string
  parent_id: number | null
  status: PageStatus
  hero_image_enabled: boolean
  hero_image?: PageUploadInput
  remove_hero_image: boolean
  seo_page_title: string
  seo_page_description: string
  page_detail?: SavePageDetailPayload
}

export type PageDocumentMultipartFile = {
  sectionIndex: number
  documentIndex: number
  file: File
}

export type PageCTABannerImageMultipartFile = {
  sectionIndex: number
  file: File
}

export type SavePageRequest = SavePagePayload & {
  heroImageFile?: File
  documentFiles?: PageDocumentMultipartFile[]
  ctaBannerImageFiles?: PageCTABannerImageMultipartFile[]
}

export type PageParentOption = {
  id: number
  page_title: string
  url_slug: string
  parent_id: number | null
}

export type PageMutationResponse = {
  message: string
  page: {
    id: number
    page_title: string
    url_slug: string
    page_type: PageType
    status: PageStatus
  }
}

function buildListQuery(filters: PageListFilters) {
  const params = new URLSearchParams()
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

function pageDocumentFileField(sectionIndex: number, documentIndex: number) {
  return `page_detail.sections[${sectionIndex}].documents.items[${documentIndex}].file`
}

function pageCTABannerImageFileField(sectionIndex: number) {
  return `page_detail.sections[${sectionIndex}].cta_banner.image.file`
}

async function fetchAllPageListItems(filters: PageListFilters) {
  const response = await apiClient.get<RawPageListResponse>(API_ROUTES.pages, {
    params: buildListQuery(filters),
  })

  return Array.isArray(response.data.items) ? response.data.items : []
}

function buildPageMultipartBody(request: SavePageRequest) {
  const { heroImageFile, documentFiles, ctaBannerImageFiles, ...payload } = request
  const files = [
    {
      fieldName: 'hero_image_file',
      file: heroImageFile,
      fileName: heroImageFile?.name,
    },
    ...(documentFiles ?? []).map((entry) => ({
      fieldName: pageDocumentFileField(entry.sectionIndex, entry.documentIndex),
      file: entry.file,
      fileName: entry.file.name,
    })),
    ...(ctaBannerImageFiles ?? []).map((entry) => ({
      fieldName: pageCTABannerImageFileField(entry.sectionIndex),
      file: entry.file,
      fileName: entry.file.name,
    })),
  ]

  return buildMultipartPayload(payload, files)
}

export function resolvePageParentId(page: PageParentRelation) {
  if (typeof page.parent_id === 'number') {
    return page.parent_id
  }

  return null
}

export function resolvePageParentSlug(page: PageParentRelation) {
  return page.parent_page_url_slug?.trim() ?? ''
}

export function isModulePage(page: { page_type?: string | null }) {
  return page.page_type === 'module'
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
    } satisfies PageListResponse
  },

  async getPage(id: number) {
    const response = await apiClient.get<PageDetailResponse>(API_ROUTES.pageById(id))
    return response.data
  },

  async listPageParentOptions() {
    const allItems = await fetchAllPageListItems({
      page: 1,
      pageSize: 100,
      searchTerm: '',
      status: '',
      sortBy: 'page_title',
      sortOrder: 'asc',
    })

    return allItems.map((item) => {
      return {
        id: item.id,
        page_title: item.page_title,
        url_slug: item.url_slug,
        parent_id: resolvePageParentId(item),
      } satisfies PageParentOption
    })
  },

  async fetchPageHeroImageContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async fetchPageCTABannerImageContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async fetchPageDocumentContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createPage(request: SavePageRequest) {
    const hasFiles =
      Boolean(request.heroImageFile) ||
      Boolean(request.documentFiles?.length) ||
      Boolean(request.ctaBannerImageFiles?.length)
    const body = hasFiles ? buildPageMultipartBody(request) : request
    const response = await apiClient.post<PageMutationResponse>(API_ROUTES.pages, body)
    return response.data
  },

  async updatePage(id: number, request: SavePageRequest) {
    const hasFiles =
      Boolean(request.heroImageFile) ||
      Boolean(request.documentFiles?.length) ||
      Boolean(request.ctaBannerImageFiles?.length)
    const body = hasFiles ? buildPageMultipartBody(request) : request
    const response = await apiClient.put<PageMutationResponse>(API_ROUTES.pageById(id), body)
    return response.data
  },

  async deletePage(id: number) {
    await apiClient.delete(API_ROUTES.pageById(id))
  },
}
