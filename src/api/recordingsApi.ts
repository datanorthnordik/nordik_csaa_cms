import { API_ROUTES } from '../constants/api'
import { apiClient } from './apiClient'

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
  storage_uri?: string | null
  gcp_object_key?: string | null
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
  storageUri?: string
  objectKey?: string
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
  recording_url?: string
  storage_uri?: string
  object_key?: string
  gcp_object_key?: string
}

export type CreateRecordingCollectionPayload = {
  name: string
  items?: RecordingItemInput[]
}

export type UpdateRecordingCollectionPayload = {
  name: string
}

export type AddRecordingItemsPayload = {
  items: RecordingItemInput[]
}

export type UpdateRecordingItemPayload = RecordingItemInput

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
  return {
    id: item.id,
    recordingCollectionId: item.recording_collection_id,
    title: item.title,
    description: item.description || '',
    recordingUrl: item.recording_url || undefined,
    storageUri: item.storage_uri || undefined,
    objectKey: item.gcp_object_key || undefined,
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

  async createRecordingCollection(payload: CreateRecordingCollectionPayload) {
    const response = await apiClient.post<RecordingCollectionMutationResponse>(
      API_ROUTES.recordings,
      payload,
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

  async addRecordingItems(id: number, payload: AddRecordingItemsPayload) {
    const response = await apiClient.post<RecordingItemsUploadResponse>(
      API_ROUTES.recordingItemsById(id),
      payload,
    )
    return response.data
  },

  async updateRecordingItem(
    collectionId: number,
    itemId: number,
    payload: UpdateRecordingItemPayload,
  ) {
    const response = await apiClient.patch<{ message: string; item: ApiRecordingItem }>(
      API_ROUTES.recordingItemById(collectionId, itemId),
      payload,
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
