import { API_ROUTES } from '../constants/api'
import type {
  MemorialCategory,
  MemorialEntry,
  MemorialEntrySummary,
  MemorialGalleryImage,
  MemorialListFilters,
  MemorialListPageMeta,
  MemorialPortraitAsset,
  MemorialStatus,
} from '../lib/memorialTypes'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

type MemorialApiSummaryEntry = {
  id: number
  full_name: string
  affiliation: string
  category: MemorialCategory
  category_label: string
  status: MemorialStatus
  date_of_birth?: string
  date_of_passing?: string
  created_at: string
  updated_at: string
  published_at?: string
}

type MemorialApiDetailResponse = MemorialApiSummaryEntry & {
  biography: string
  portrait?: {
    file_name: string
    mime_type: string
    file_size: number
  }
  gallery_images: Array<{
    id: number
    file_name: string
    mime_type: string
    file_size: number
  }>
}

type MemorialApiListResponse = {
  items: MemorialApiSummaryEntry[]
  pagination: {
    page: number
    page_size: number
    total_items: number
    total_pages: number
    has_next: boolean
    has_prev: boolean
  }
  applied_filters: {
    page: number
    page_size: number
    search_term: string
    status: MemorialListFilters['status']
    category: string
  }
}

type MemorialApiMutationResponse = {
  message: string
  memorial: {
    id: number
    full_name: string
    category: MemorialCategory
    status: MemorialStatus
    updated_at: string
  }
}

export type MemorialApiMutationInput = {
  full_name: string
  affiliation: string
  category: MemorialCategory
  status: MemorialStatus
  biography: string
  date_of_birth?: string
  date_of_passing?: string
  remove_portrait?: boolean
  remove_gallery_image_ids?: number[]
}

function buildListQuery(filters: MemorialListFilters) {
  const params = new URLSearchParams()
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
  params.set('status', filters.status)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }
  if (filters.category) {
    params.set('category', filters.category)
  }

  return params
}

function memorialApiSummaryToLocal(entry: MemorialApiSummaryEntry): MemorialEntrySummary {
  return {
    id: String(entry.id),
    fullName: entry.full_name,
    affiliation: entry.affiliation,
    category: entry.category,
    categoryLabel: entry.category_label,
    status: entry.status,
    dateOfBirth: entry.date_of_birth ?? '',
    dateOfPassing: entry.date_of_passing ?? '',
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
    publishedAt: entry.published_at,
  }
}

function memorialPortraitToLocal(
  portrait?: MemorialApiDetailResponse['portrait'],
): MemorialPortraitAsset | undefined {
  if (!portrait) {
    return undefined
  }

  return {
    fileName: portrait.file_name,
    mimeType: portrait.mime_type,
    fileSize: portrait.file_size,
  }
}

function memorialGalleryImageToLocal(
  image: MemorialApiDetailResponse['gallery_images'][number],
): MemorialGalleryImage {
  return {
    id: String(image.id),
    fileName: image.file_name,
    mimeType: image.mime_type,
    fileSize: image.file_size,
  }
}

function memorialApiDetailToLocal(entry: MemorialApiDetailResponse): MemorialEntry {
  return {
    ...memorialApiSummaryToLocal(entry),
    biography: entry.biography,
    portrait: memorialPortraitToLocal(entry.portrait),
    galleryImages: entry.gallery_images.map(memorialGalleryImageToLocal),
  }
}

function memorialApiPageToLocal(
  pagination: MemorialApiListResponse['pagination'],
): MemorialListPageMeta {
  return {
    page: pagination.page,
    pageSize: pagination.page_size,
    totalItems: pagination.total_items,
    totalPages: pagination.total_pages,
    hasNext: pagination.has_next,
    hasPrev: pagination.has_prev,
  }
}

function buildMutationBody(
  input: MemorialApiMutationInput,
  portraitFile?: File | null,
  galleryFiles: File[] = [],
) {
  const payload = {
    ...input,
    date_of_birth: input.date_of_birth || '',
    date_of_passing: input.date_of_passing || '',
    remove_portrait: Boolean(input.remove_portrait),
    remove_gallery_image_ids: input.remove_gallery_image_ids ?? [],
    portrait: portraitFile
      ? {
          file_name: portraitFile.name,
          mime_type: portraitFile.type,
          file_size: portraitFile.size,
        }
      : undefined,
    gallery_images: galleryFiles.map((file) => ({
      file_name: file.name,
      mime_type: file.type,
      file_size: file.size,
    })),
  }

  if (!portraitFile && galleryFiles.length === 0) {
    return payload
  }

  const files = [
    portraitFile
      ? {
          fieldName: 'portrait_file',
          file: portraitFile,
          fileName: portraitFile.name,
        }
      : null,
    ...galleryFiles.map((file, index) => ({
      fieldName: `gallery_images[${index}].file`,
      file,
      fileName: file.name,
    })),
  ].filter(Boolean)

  return buildMultipartPayload(
    payload,
    files as Array<{ fieldName: string; file?: Blob | null; fileName?: string }>,
  )
}

export const memorialApi = {
  async listMemorials(filters: MemorialListFilters) {
    const response = await apiClient.get<MemorialApiListResponse>(API_ROUTES.memorial, {
      params: buildListQuery(filters),
    })

    return {
      items: response.data.items.map(memorialApiSummaryToLocal),
      pagination: memorialApiPageToLocal(response.data.pagination),
    }
  },

  async getMemorial(id: string) {
    const response = await apiClient.get<MemorialApiDetailResponse>(API_ROUTES.memorialById(id))
    return memorialApiDetailToLocal(response.data)
  },

  async getMemorialPortraitContent(id: string) {
    const response = await apiClient.get<Blob>(API_ROUTES.memorialPortraitById(id), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async getMemorialGalleryImageContent(id: string, imageId: string) {
    const response = await apiClient.get<Blob>(
      API_ROUTES.memorialGalleryImageContentById(id, imageId),
      {
        responseType: 'blob',
        skipErrorToast: true,
      },
    )
    return response.data
  },

  async createMemorial(
    input: MemorialApiMutationInput,
    portraitFile?: File | null,
    galleryFiles: File[] = [],
  ) {
    const response = await apiClient.post<MemorialApiMutationResponse>(
      API_ROUTES.memorial,
      buildMutationBody(input, portraitFile, galleryFiles),
    )
    return response.data.memorial
  },

  async updateMemorial(
    id: string,
    input: MemorialApiMutationInput,
    portraitFile?: File | null,
    galleryFiles: File[] = [],
  ) {
    const response = await apiClient.put<MemorialApiMutationResponse>(
      API_ROUTES.memorialById(id),
      buildMutationBody(input, portraitFile, galleryFiles),
    )
    return response.data.memorial
  },

  async deleteMemorial(id: string) {
    await apiClient.delete(API_ROUTES.memorialById(id))
  },
}
