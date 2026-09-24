import { API_BASE_URL, API_ROUTES } from '../constants/api'
import { assertValidRecordingUploadFile } from '../lib/recordingUpload'
import { apiClient } from './apiClient'
import { buildMultipartPayload } from './multipartForm'

type ApiRecordingCollectionSummary = {
  id: number
  name: string
  item_count: number
  created_at: string
  updated_at: string
}

type ApiRecordingItem = {
  id: number
  recording_collection_id: number
  title: string
  description: string | null
  recording_url?: string | null
  sort_order: number
  created_at: string
  updated_at: string
}

type ApiRecordingCollectionDetail = {
  id: number
  name: string
  item_count: number
  items: ApiRecordingItem[] | null
  created_at: string
  updated_at: string
}

type RecordingCollectionListResponse = {
  items: ApiRecordingCollectionSummary[]
}

export type RecordingCollectionSummary = {
  id: number
  name: string
  itemCount: number
  createdAt: string
  updatedAt: string
}

export type RecordingItem = {
  id: number
  recordingCollectionId: number
  title: string
  description: string
  recordingUrl?: string
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type RecordingCollectionDetail = {
  id: number
  name: string
  itemCount: number
  items: RecordingItem[]
  createdAt: string
  updatedAt: string
}

export type RecordingItemInput = {
  title: string
  description?: string
  file_name?: string
  mime_type?: string
  recording_url?: string
  storage_uri?: string
  object_key?: string
  gcp_object_key?: string
}

export type CreateRecordingCollectionPayload = {
  name: string
  items?: RecordingItemInput[]
}

export type CreateRecordingCollectionRequest = CreateRecordingCollectionPayload & {
  recordingFiles?: Array<File | null | undefined>
}

export type UpdateRecordingCollectionPayload = {
  name: string
}

export type AddRecordingItemsPayload = {
  items: RecordingItemInput[]
}

export type AddRecordingItemsRequest = AddRecordingItemsPayload & {
  recordingFiles?: Array<File | null | undefined>
}

export type UpdateRecordingItemPayload = RecordingItemInput

export type UpdateRecordingItemRequest = UpdateRecordingItemPayload & {
  recordingFile?: File | null
}

export type RecordingCollectionMutationResponse = {
  message: string
  recording: {
    id: number
    name: string
  }
}

export type RecordingItemsUploadResponse = {
  message: string
  uploadedCount: number
}

export type RecordingItemMutationResponse = {
  message: string
  item: RecordingItem
}

export type DeleteRecordingCollectionResponse = {
  message: string
}

export type DeleteRecordingItemResponse = {
  message: string
  deletedCount: number
}

function mapRecordingCollectionSummary(
  item: ApiRecordingCollectionSummary,
): RecordingCollectionSummary {
  return {
    id: item.id,
    name: item.name,
    itemCount: item.item_count,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapRecordingItem(item: ApiRecordingItem): RecordingItem {
  const recordingUrl = item.recording_url?.trim()
  return {
    id: item.id,
    recordingCollectionId: item.recording_collection_id,
    title: item.title,
    description: item.description || '',
    recordingUrl: recordingUrl
      ? new URL(recordingUrl.replace(/^\/+/, ''), `${API_BASE_URL}/`).toString()
      : undefined,
    sortOrder: item.sort_order,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapRecordingCollectionDetail(
  detail: ApiRecordingCollectionDetail,
): RecordingCollectionDetail {
  return {
    id: detail.id,
    name: detail.name,
    itemCount: detail.item_count,
    items: (detail.items ?? []).map(mapRecordingItem),
    createdAt: detail.created_at,
    updatedAt: detail.updated_at,
  }
}

function recordingItemFileField(index: number) {
  return `items[${index}].recording_file`
}

function buildItemsBody<T extends { recordingFiles?: Array<File | null | undefined> }>(
  request: T,
) {
  const { recordingFiles, ...payload } = request
  recordingFiles?.forEach((file) => {
    if (file) {
      assertValidRecordingUploadFile(file)
    }
  })

  if (!recordingFiles?.some(Boolean)) {
    return payload
  }

  return buildMultipartPayload(
    payload,
    recordingFiles.map((file, index) => ({
      fieldName: recordingItemFileField(index),
      file,
      fileName: file?.name,
    })),
  )
}

function buildUpdateItemBody(request: UpdateRecordingItemRequest) {
  const { recordingFile, ...payload } = request
  if (!recordingFile) {
    return payload
  }

  assertValidRecordingUploadFile(recordingFile)
  return buildMultipartPayload(payload, [
    {
      fieldName: 'recording_file',
      file: recordingFile,
      fileName: recordingFile.name,
    },
  ])
}

export const recordingsApi = {
  async listRecordingCollections() {
    const response = await apiClient.get<RecordingCollectionListResponse>(API_ROUTES.recordings)
    return (response.data.items ?? []).map(mapRecordingCollectionSummary)
  },

  async getRecordingCollection(id: number) {
    const response = await apiClient.get<ApiRecordingCollectionDetail>(
      API_ROUTES.recordingById(id),
    )
    return mapRecordingCollectionDetail(response.data)
  },

  async createRecordingCollection(request: CreateRecordingCollectionRequest) {
    const response = await apiClient.post<RecordingCollectionMutationResponse>(
      API_ROUTES.recordings,
      buildItemsBody(request),
    )
    return response.data
  },

  async updateRecordingCollection(id: number, payload: UpdateRecordingCollectionPayload) {
    const response = await apiClient.put<RecordingCollectionMutationResponse>(
      API_ROUTES.recordingById(id),
      payload,
    )
    return response.data
  },

  async deleteRecordingCollection(id: number) {
    const response = await apiClient.delete<DeleteRecordingCollectionResponse>(
      API_ROUTES.recordingById(id),
    )
    return response.data
  },

  async addRecordingItems(id: number, request: AddRecordingItemsRequest) {
    const response = await apiClient.post<RecordingItemsUploadResponse>(
      API_ROUTES.recordingItemsById(id),
      buildItemsBody(request),
    )
    return response.data
  },

  async updateRecordingItem(
    collectionId: number,
    itemId: number,
    request: UpdateRecordingItemRequest,
  ) {
    const response = await apiClient.patch<{ message: string; item: ApiRecordingItem }>(
      API_ROUTES.recordingItemById(collectionId, itemId),
      buildUpdateItemBody(request),
    )

    return {
      message: response.data.message,
      item: mapRecordingItem(response.data.item),
    } satisfies RecordingItemMutationResponse
  },

  async deleteRecordingItem(collectionId: number, itemId: number) {
    const response = await apiClient.delete<DeleteRecordingItemResponse>(
      API_ROUTES.recordingItemById(collectionId, itemId),
    )
    return response.data
  },
}
