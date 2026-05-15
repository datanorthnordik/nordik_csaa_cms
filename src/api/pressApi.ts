import { API_ROUTES } from '../constants/api'
import type { PressEntry, PressMedia, PressStatus, PressVisibility } from '../lib/pressTypes'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type PressApiEntry = {
  id: number
  title: string
  release_date: string
  category_id: number | null
  source_url: string
  content_html: string
  status: PressStatus
  visibility: PressVisibility
  publish_at: string | null
  cover_image_url?: string
  created_at: string
  updated_at: string
}

export type PressApiMedia = {
  id: number
  display_name: string
  file_name: string
  gcp_object_key?: string
  file_url: string
  mime_type: string
  file_size: number
  media_role: 'attachment' | 'cover_image'
  sort_order: number
  created_at: string
  updated_at: string
}

export type PressApiDetailEntry = PressApiEntry & {
  media: PressApiMedia[]
}

export type PressApiListResponse = {
  items: PressApiEntry[]
  total: number
  page: number
  page_size: number
  total_pages: number
}

export type PressApiCreateInput = {
  title: string
  release_date: string
  category_id: number | null
  source_url: string
  content_html: string
  status: PressStatus
  visibility: PressVisibility
  publish_at: string | null
  remove_cover_image?: boolean
}

export type PressMutationResponse = {
  message: string
  entry: {
    id: number
    title: string
    release_date: string
    status: PressStatus
    visibility: PressVisibility
  }
}

export type PressMediaUploadInput = {
  display_name: string
  file_name?: string
}

export type AddPressMediaResponse = {
  message: string
  uploadedCount: number
}

export type UpdatePressMediaInput = {
  display_name?: string
  file_name?: string
}

export type UpdatePressMediaResponse = {
  message: string
  media: PressApiMedia
}

export type DeletePressMediaRequest = {
  media_ids: number[]
}

export type DeletePressMediaResponse = {
  message: string
  deletedCount: number
}

export type ReorderPressMediaRequest = {
  media_ids: number[]
}

export type ReorderPressMediaResponse = {
  message: string
  updatedCount: number
}

function parseCategoryId(categoryId: PressEntry['categoryId']) {
  if (!categoryId) {
    return null
  }

  const parsed = Number.parseInt(categoryId, 10)
  return Number.isNaN(parsed) ? null : parsed
}

function normalizeDateOnly(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return ''
  }

  const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/)
  if (match) {
    return match[1]
  }

  const parsed = Date.parse(trimmed)
  if (Number.isNaN(parsed)) {
    return trimmed
  }

  return new Date(parsed).toISOString().slice(0, 10)
}

export function pressApiMediaToLocal(media: PressApiMedia): PressMedia {
  return {
    id: String(media.id),
    fileName: media.file_name || media.display_name || `press-media-${media.id}`,
    fileUrl: media.file_url,
    mimeType: media.mime_type,
    fileSize: media.file_size,
  }
}

export function pressApiEntryToLocal(entry: PressApiEntry): PressEntry {
  return {
    id: String(entry.id),
    title: entry.title,
    releaseDate: normalizeDateOnly(entry.release_date),
    categoryId: entry.category_id != null ? String(entry.category_id) : null,
    sourceUrl: entry.source_url,
    contentHtml: entry.content_html,
    status: entry.status,
    visibility: entry.visibility,
    publishAt: entry.publish_at,
    coverImageUrl: entry.cover_image_url,
    media: [],
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

export function pressApiDetailToLocal(entry: PressApiDetailEntry): PressEntry {
  return {
    ...pressApiEntryToLocal(entry),
    media: (entry.media ?? []).map(pressApiMediaToLocal),
  }
}

export function localEntryToPressApiInput(entry: PressEntry): PressApiCreateInput {
  return {
    title: entry.title,
    release_date: entry.releaseDate,
    category_id: parseCategoryId(entry.categoryId),
    source_url: entry.sourceUrl,
    content_html: entry.contentHtml,
    status: entry.status,
    visibility: entry.visibility,
    publish_at: entry.publishAt,
  }
}

export async function fetchPressEntries(
  params?: Record<string, string | number>,
): Promise<PressApiListResponse> {
  const { data } = await apiClient.get<PressApiListResponse>(API_ROUTES.press, { params })
  return data
}

export async function fetchPressEntry(id: string): Promise<PressEntry> {
  const { data } = await apiClient.get<PressApiDetailEntry>(API_ROUTES.pressById(id))
  return pressApiDetailToLocal(data)
}

export async function fetchPressCoverImageContent(entryId: string): Promise<Blob> {
  const response = await apiClient.get<Blob>(API_ROUTES.pressCoverById(entryId), {
    responseType: 'blob',
    skipErrorToast: true,
  })

  return response.data
}

export async function createPressApiEntry(
  input: PressApiCreateInput,
  coverImageFile?: File,
): Promise<PressMutationResponse> {
  const body = coverImageFile
    ? buildMultipartPayload(input, [
        {
          fieldName: 'cover_image_file',
          file: coverImageFile,
          fileName: coverImageFile.name,
        },
      ])
    : input

  const { data } = await apiClient.post<PressMutationResponse>(API_ROUTES.press, body)
  return data
}

export async function updatePressApiEntry(
  id: string,
  patch: Partial<PressApiCreateInput>,
  coverImageFile?: File,
): Promise<PressMutationResponse> {
  const body = coverImageFile
    ? buildMultipartPayload(patch, [
        {
          fieldName: 'cover_image_file',
          file: coverImageFile,
          fileName: coverImageFile.name,
        },
      ])
    : patch

  const { data } = await apiClient.put<PressMutationResponse>(API_ROUTES.pressById(id), body)
  return data
}

export async function deletePressApiEntry(id: string): Promise<void> {
  await apiClient.delete(API_ROUTES.pressById(id))
}

export async function addPressMedia(
  entryId: string,
  files: File[],
  metadata: PressMediaUploadInput[],
): Promise<AddPressMediaResponse> {
  const media = files.map((file, index) => ({
    display_name: metadata[index]?.display_name || file.name,
    file_name: metadata[index]?.file_name || file.name,
  }))

  const body = buildMultipartPayload(
    { media },
    files.map((file, index) => ({
      fieldName: `media[${index}].file`,
      file,
      fileName: file.name,
    })),
  )

  const { data } = await apiClient.post<AddPressMediaResponse>(
    `${API_ROUTES.pressById(entryId)}/media`,
    body,
  )

  return data
}

export async function updatePressMedia(
  entryId: string,
  mediaId: string,
  input: UpdatePressMediaInput,
): Promise<PressMedia> {
  const { data } = await apiClient.patch<UpdatePressMediaResponse>(
    `${API_ROUTES.pressById(entryId)}/media/${mediaId}`,
    input,
  )

  return pressApiMediaToLocal(data.media)
}

export async function getPressMediaContent(entryId: string, mediaId: string): Promise<Blob> {
  const response = await apiClient.get<Blob>(
    `${API_ROUTES.pressById(entryId)}/media/${mediaId}/content`,
    { responseType: 'blob' },
  )

  return response.data
}

export async function reorderPressMedia(
  entryId: string,
  mediaIds: number[],
): Promise<ReorderPressMediaResponse> {
  const { data } = await apiClient.put<ReorderPressMediaResponse>(
    `${API_ROUTES.pressById(entryId)}/media/order`,
    { media_ids: mediaIds } satisfies ReorderPressMediaRequest,
  )

  return data
}

export async function deletePressMedia(
  entryId: string,
  mediaIds: number[],
): Promise<DeletePressMediaResponse> {
  const { data } = await apiClient.delete<DeletePressMediaResponse>(
    `${API_ROUTES.pressById(entryId)}/media`,
    { data: { media_ids: mediaIds } satisfies DeletePressMediaRequest },
  )

  return data
}
