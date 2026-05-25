import { API_ROUTES } from '../constants/api'
import {
  type ResourceCategoryCount,
  type ResourceEntry,
  type ResourceFileType,
  type ResourceListFilters,
  type ResourceListPageMeta,
  type ResourceVisibility,
} from '../lib/resourceTypes'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type ResourceApiEntry = {
  id: number
  name: string
  category: ResourceEntry['category']
  category_label: string
  visibility: ResourceVisibility
  file_name: string
  mime_type: string
  file_size: number
  content_url: string
  created_at: string
  updated_at: string
}

export type ResourceApiListResponse = {
  items: ResourceApiEntry[]
  pagination: {
    page: number
    page_size: number
    total_items: number
    total_pages: number
    has_next: boolean
    has_prev: boolean
  }
  summary: {
    category_counts: Array<{
      category: ResourceCategoryCount['category']
      label: string
      count: number
    }>
  }
  applied_filters: {
    page: number
    page_size: number
    search_term: string
    category: string
    file_type: ResourceFileType
  }
}

export type ResourceApiMutationInput = {
  name: string
  category: ResourceEntry['category']
  visibility: ResourceVisibility
}

export type ResourceApiMutationResponse = {
  message: string
  resource: {
    id: number
    name: string
    category: ResourceEntry['category']
    visibility: ResourceVisibility
    updated_at: string
  }
}

function buildListQuery(filters: ResourceListFilters) {
  const params = new URLSearchParams()
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
  params.set('file_type', filters.fileType)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }
  if (filters.category) {
    params.set('category', filters.category)
  }

  return params
}

export function resourceApiEntryToLocal(entry: ResourceApiEntry): ResourceEntry {
  return {
    id: String(entry.id),
    name: entry.name,
    category: entry.category,
    categoryLabel: entry.category_label,
    visibility: entry.visibility,
    fileName: entry.file_name,
    mimeType: entry.mime_type,
    fileSize: entry.file_size,
    contentUrl: entry.content_url,
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

export function resourceApiPageToLocal(pagination: ResourceApiListResponse['pagination']): ResourceListPageMeta {
  return {
    page: pagination.page,
    pageSize: pagination.page_size,
    totalItems: pagination.total_items,
    totalPages: pagination.total_pages,
    hasNext: pagination.has_next,
    hasPrev: pagination.has_prev,
  }
}

export const resourcesApi = {
  async listResources(filters: ResourceListFilters) {
    const response = await apiClient.get<ResourceApiListResponse>(API_ROUTES.resources, {
      params: buildListQuery(filters),
    })

    return {
      items: response.data.items.map(resourceApiEntryToLocal),
      pagination: resourceApiPageToLocal(response.data.pagination),
      summary: {
        categoryCounts: response.data.summary.category_counts,
      },
      appliedFilters: {
        page: response.data.applied_filters.page,
        pageSize: response.data.applied_filters.page_size,
        searchTerm: response.data.applied_filters.search_term,
        category: response.data.applied_filters.category as ResourceListFilters['category'],
        fileType: response.data.applied_filters.file_type,
      },
    }
  },

  async getResource(id: string) {
    const response = await apiClient.get<ResourceApiEntry>(API_ROUTES.resourceById(id))
    return resourceApiEntryToLocal(response.data)
  },

  async getResourceContent(id: string) {
    const response = await apiClient.get<Blob>(API_ROUTES.resourceContentById(id), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createResource(input: ResourceApiMutationInput, file: File) {
    const body = buildMultipartPayload(input, [
      {
        fieldName: 'resource_file',
        file,
        fileName: file.name,
      },
    ])

    const response = await apiClient.post<ResourceApiMutationResponse>(API_ROUTES.resources, body)
    return response.data.resource
  },

  async updateResource(id: string, input: ResourceApiMutationInput, file?: File) {
    const body = file
      ? buildMultipartPayload(input, [
          {
            fieldName: 'resource_file',
            file,
            fileName: file.name,
          },
        ])
      : input

    const response = await apiClient.put<ResourceApiMutationResponse>(
      API_ROUTES.resourceById(id),
      body,
    )
    return response.data.resource
  },

  async deleteResource(id: string) {
    await apiClient.delete(API_ROUTES.resourceById(id))
  },
}
