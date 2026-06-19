import { API_BASE_URL, API_ROUTES } from '../constants/api'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type BlogSectionType =
  | 'heading'
  | 'image'
  | 'typography'
  | 'action'
  | 'video'
  | 'animation'

export type BlogActionType = 'link' | 'video'
export type BlogAnimationNavigation = 'vertical' | 'horizontal'
export type BlogAnimationImagePosition = 'left' | 'right'

type JSONObject = Record<string, unknown>

export type BlogListItem = {
  id: number
  publish_date: string
  heading: string
  description: string
  cover_image_url: string
  cover_image_object_key: string
  cover_image_fetch_url: string
  updated_by?: number | null
  updated_by_name: string
  created_at: string
  updated_at: string
}

export type BlogListFilters = {
  page: number
  pageSize: number
  searchTerm: string
  sortBy: 'publish_date' | 'heading' | 'updated_at' | 'created_at'
  sortOrder: 'asc' | 'desc'
}

export type BlogSectionAssetResponse = {
  file_url: string
  fetch_url: string
  storage_uri: string
  gcp_object_key: string
}

export type BlogHeadingSectionResponse = {
  heading_text: string
  underline_enabled: boolean
}

export type BlogImageSectionResponse = {
  asset?: BlogSectionAssetResponse | null
  caption: string
}

export type BlogTypographySectionResponse = {
  html_content: string
  text_content: string
}

export type BlogActionSectionResponse = {
  text: string
  action_type: BlogActionType
  target_url: string
}

export type BlogVideoSectionResponse = {
  youtube_url: string
  caption: string
}

export type BlogAnimationItemResponse = {
  id: number
  sort_order: number
  heading: string
  sub_heading: string
  description: string
  image?: BlogSectionAssetResponse | null
}

export type BlogAnimationSectionResponse = {
  navigation: BlogAnimationNavigation
  image_position: BlogAnimationImagePosition
  items: BlogAnimationItemResponse[]
}

export type BlogSectionResponse = {
  id: number
  section_name: string
  section_type: BlogSectionType
  sort_order: number
  is_enabled: boolean
  settings?: JSONObject | null
  heading?: BlogHeadingSectionResponse | null
  image?: BlogImageSectionResponse | null
  typography?: BlogTypographySectionResponse | null
  action?: BlogActionSectionResponse | null
  video?: BlogVideoSectionResponse | null
  animation?: BlogAnimationSectionResponse | null
}

export type BlogDetailResponse = {
  id: number
  publish_date: string
  heading: string
  description: string
  cover_image_url: string
  cover_image_object_key: string
  cover_image_fetch_url: string
  created_by?: number | null
  created_by_name: string
  updated_by?: number | null
  updated_by_name: string
  created_at: string
  updated_at: string
  blog_detail?: {
    sections: BlogSectionResponse[]
  } | null
}

export type BlogUploadInput = {
  file_name?: string
  mime_type?: string
  file_url?: string
  storage_uri?: string
  object_key?: string
  gcp_object_key?: string
}

export type SaveBlogHeadingSectionPayload = {
  heading_text: string
  underline_enabled: boolean
}

export type SaveBlogImageSectionPayload = {
  asset?: BlogUploadInput
  caption: string
}

export type SaveBlogTypographySectionPayload = {
  html_content: string
  text_content: string
}

export type SaveBlogActionSectionPayload = {
  text: string
  action_type: BlogActionType
  target_url: string
}

export type SaveBlogVideoSectionPayload = {
  youtube_url: string
  caption: string
}

export type SaveBlogAnimationItemPayload = {
  id?: number
  sort_order: number
  heading: string
  sub_heading: string
  description: string
  image?: BlogUploadInput
}

export type SaveBlogAnimationSectionPayload = {
  navigation: BlogAnimationNavigation
  image_position: BlogAnimationImagePosition
  items: SaveBlogAnimationItemPayload[]
}

export type SaveBlogSectionPayload = {
  id?: number
  section_name: string
  section_type: BlogSectionType
  sort_order: number
  is_enabled: boolean
  settings?: JSONObject
  heading?: SaveBlogHeadingSectionPayload
  image?: SaveBlogImageSectionPayload
  typography?: SaveBlogTypographySectionPayload
  action?: SaveBlogActionSectionPayload
  video?: SaveBlogVideoSectionPayload
  animation?: SaveBlogAnimationSectionPayload
}

export type SaveBlogPayload = {
  publish_date: string
  heading: string
  description: string
  cover_image?: BlogUploadInput
  remove_cover_image: boolean
  blog_detail?: {
    sections: SaveBlogSectionPayload[]
  }
}

export type BlogSectionImageMultipartFile = {
  sectionIndex: number
  file: File
}

export type BlogAnimationItemImageMultipartFile = {
  sectionIndex: number
  itemIndex: number
  file: File
}

export type SaveBlogRequest = SaveBlogPayload & {
  coverImageFile?: File
  sectionImageFiles?: BlogSectionImageMultipartFile[]
  animationItemImageFiles?: BlogAnimationItemImageMultipartFile[]
}

export type BlogMutationResponse = {
  message: string
  blog: {
    id: number
    publish_date: string
    heading: string
  }
}

function buildListQuery(filters: BlogListFilters) {
  const params = new URLSearchParams()
  params.set('sort_by', filters.sortBy)
  params.set('sort_order', filters.sortOrder)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }

  return params
}

function sectionImageFileField(sectionIndex: number) {
  return `blog_detail.sections[${sectionIndex}].image.asset.file`
}

function animationItemImageField(sectionIndex: number, itemIndex: number) {
  return `blog_detail.sections[${sectionIndex}].animation.items[${itemIndex}].image.file`
}

function buildBlogMultipartBody(request: SaveBlogRequest) {
  const { coverImageFile, sectionImageFiles, animationItemImageFiles, ...payload } = request

  const files = [
    {
      fieldName: 'cover_image_file',
      file: coverImageFile,
      fileName: coverImageFile?.name,
    },
    ...(sectionImageFiles ?? []).map((entry) => ({
      fieldName: sectionImageFileField(entry.sectionIndex),
      file: entry.file,
      fileName: entry.file.name,
    })),
    ...(animationItemImageFiles ?? []).map((entry) => ({
      fieldName: animationItemImageField(entry.sectionIndex, entry.itemIndex),
      file: entry.file,
      fileName: entry.file.name,
    })),
  ]

  return buildMultipartPayload(payload, files)
}

export function resolveBlogAssetUrl(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return ''
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed) || trimmed.startsWith('//')) {
    return trimmed
  }
  try {
    return new URL(trimmed, `${API_BASE_URL}/`).toString()
  } catch {
    return trimmed
  }
}

export const blogsApi = {
  async listBlogs(filters: BlogListFilters) {
    const response = await apiClient.get<{
      items: BlogListItem[] | null
    }>(API_ROUTES.blogs, {
      params: buildListQuery(filters),
    })

    return Array.isArray(response.data.items) ? response.data.items : []
  },

  async getBlog(id: number) {
    const response = await apiClient.get<BlogDetailResponse>(API_ROUTES.blogById(id))
    return response.data
  },

  async createBlog(request: SaveBlogRequest) {
    const hasFiles =
      Boolean(request.coverImageFile) ||
      Boolean(request.sectionImageFiles?.length) ||
      Boolean(request.animationItemImageFiles?.length)
    const body = hasFiles ? buildBlogMultipartBody(request) : request
    const response = await apiClient.post<BlogMutationResponse>(API_ROUTES.blogs, body)
    return response.data
  },

  async updateBlog(id: number, request: SaveBlogRequest) {
    const hasFiles =
      Boolean(request.coverImageFile) ||
      Boolean(request.sectionImageFiles?.length) ||
      Boolean(request.animationItemImageFiles?.length)
    const body = hasFiles ? buildBlogMultipartBody(request) : request
    const response = await apiClient.put<BlogMutationResponse>(API_ROUTES.blogById(id), body)
    return response.data
  },

  async deleteBlog(id: number) {
    await apiClient.delete(API_ROUTES.blogById(id))
  },
}
