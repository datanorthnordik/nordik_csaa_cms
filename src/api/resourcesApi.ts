import { API_ROUTES } from '../constants/api'
import { assertValidResourceUploadFile } from '../lib/resourceUpload'
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
  description: string
  category: ResourceEntry['category']
  category_label: string
  visibility: ResourceVisibility
  link_url: string
  file_name: string
  mime_type: string
  file_size: number
  has_document: boolean
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
  description: string
  category: ResourceEntry['category']
  visibility: ResourceVisibility
  linkUrl?: string
}

export type ResourceApiMutationResponse = {
  message: string
  resource: {
    id: number
    name: string
    description: string
    category: ResourceEntry['category']
    visibility: ResourceVisibility
    link_url: string
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

function resourceMutationInputToApi(input: ResourceApiMutationInput) {
  return {
    name: input.name,
    description: input.description,
    category: input.category,
    visibility: input.visibility,
    link_url: input.category === 'link' ? input.linkUrl ?? '' : '',
  }
}

function buildResourceMutationBody(input: ResourceApiMutationInput, file?: File) {
  const payload = resourceMutationInputToApi(input)

  if (!file) {
    return payload
  }

  assertValidResourceUploadFile(file)

  return buildMultipartPayload(payload, [
    {
      fieldName: 'resource_file',
      file,
      fileName: file.name,
    },
  ])
}

export function resourceApiEntryToLocal(entry: ResourceApiEntry): ResourceEntry {
  return {
    id: String(entry.id),
    name: entry.name,
    description: entry.description,
    category: entry.category,
    categoryLabel: entry.category_label,
    visibility: entry.visibility,
    linkUrl: entry.link_url,
    fileName: entry.file_name,
    mimeType: entry.mime_type,
    fileSize: entry.file_size,
    hasDocument: entry.has_document,
    contentUrl: entry.content_url,
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

export function resourceApiPageToLocal(
  pagination: ResourceApiListResponse['pagination'],
): ResourceListPageMeta {
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
    const response = await apiClient.get<ResourceApiListResponse>(
      API_ROUTES.resources,
      {
        params: buildListQuery(filters),
      },
    )

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
    const response = await apiClient.get<ResourceApiEntry>(
      API_ROUTES.resourceById(id),
    )

    return resourceApiEntryToLocal(response.data)
  },

  async getResourceContent(id: string) {
    const response = await apiClient.get<Blob>(
      API_ROUTES.resourceContentById(id),
      {
        responseType: 'blob',
        skipErrorToast: true,
      },
    )

    return response.data
  },

  async createResource(input: ResourceApiMutationInput, file?: File) {
    const body = buildResourceMutationBody(input, file)

    const response = await apiClient.post<ResourceApiMutationResponse>(
      API_ROUTES.resources,
      body,
    )

    return response.data.resource
  },

  async updateResource(id: string, input: ResourceApiMutationInput, file?: File) {
    const body = buildResourceMutationBody(input, file)

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
