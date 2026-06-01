import { API_ROUTES } from '../constants/api'
import type {
  NewsletterEntry,
  NewsletterMedia,
  NewsletterStatus,
  NewsletterVisibility,
} from '../lib/newsletterTypes'
import { assertValidResourceUploadFile } from '../lib/resourceUpload'
import { apiClient } from './apiClient'
import { buildMultipartPayload } from './multipartForm'

export type NewsletterApiEntry = {
  id: number
  title: string
  category: string
  send_date: string
  content_html: string
  status: NewsletterStatus
  visibility: NewsletterVisibility
  publish_at: string | null
  created_at: string
  updated_at: string
}

export type NewsletterApiMedia = {
  id: number
  display_name: string
  file_name: string
  gcp_object_key?: string
  file_url: string
  mime_type: string
  file_size: number
  media_role: 'attachment'
  sort_order: number
  created_at: string
  updated_at: string
}

export type NewsletterApiDetailEntry = NewsletterApiEntry & {
  media: NewsletterApiMedia[]
}

export type NewsletterApiListResponse = {
  items: NewsletterApiEntry[]
  total: number
  page: number
  page_size: number
  total_pages: number
}

export type NewsletterApiCreateInput = {
  title: string
  category: string
  send_date: string
  content_html: string
  status: NewsletterStatus
  visibility: NewsletterVisibility
  publish_at: string | null
}

export type NewsletterMutationResponse = {
  message: string
  entry: {
    id: number
    title: string
    category: string
    send_date: string
    status: NewsletterStatus
    visibility: NewsletterVisibility
  }
}

export type NewsletterMediaUploadInput = {
  display_name: string
  file_name?: string
}

export type AddNewsletterMediaResponse = {
  message: string
  uploadedCount: number
}

export type UpdateNewsletterMediaInput = {
  display_name?: string
  file_name?: string
}

export type UpdateNewsletterMediaResponse = {
  message: string
  media: NewsletterApiMedia
}

export type DeleteNewsletterMediaRequest = {
  media_ids: number[]
}

export type DeleteNewsletterMediaResponse = {
  message: string
  deletedCount: number
}

export type ReorderNewsletterMediaRequest = {
  media_ids: number[]
}

export type ReorderNewsletterMediaResponse = {
  message: string
  updatedCount: number
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

export function newsletterApiMediaToLocal(media: NewsletterApiMedia): NewsletterMedia {
  return {
    id: String(media.id),
    fileName: media.file_name || media.display_name || `newsletter-media-${media.id}`,
    fileUrl: media.file_url,
    mimeType: media.mime_type,
    fileSize: media.file_size,
  }
}

export function newsletterApiEntryToLocal(entry: NewsletterApiEntry): NewsletterEntry {
  return {
    id: String(entry.id),
    title: entry.title,
    category: entry.category === 'csaa' || entry.category === 'cst' ? entry.category : '',
    sendDate: normalizeDateOnly(entry.send_date),
    contentHtml: entry.content_html,
    status: entry.status,
    visibility: entry.visibility,
    publishAt: entry.publish_at,
    media: [],
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

export function newsletterApiDetailToLocal(entry: NewsletterApiDetailEntry): NewsletterEntry {
  return {
    ...newsletterApiEntryToLocal(entry),
    media: (entry.media ?? []).map(newsletterApiMediaToLocal),
  }
}

export async function fetchNewsletterEntries(
  params?: Record<string, string | number>,
): Promise<NewsletterApiListResponse> {
  const { data } = await apiClient.get<NewsletterApiListResponse>(API_ROUTES.newsletters, {
    params,
  })
  return data
}

export async function fetchNewsletterEntry(id: string): Promise<NewsletterEntry> {
  const { data } = await apiClient.get<NewsletterApiDetailEntry>(
    API_ROUTES.newsletterById(id),
  )
  return newsletterApiDetailToLocal(data)
}

export async function createNewsletterApiEntry(
  input: NewsletterApiCreateInput,
): Promise<NewsletterMutationResponse> {
  const { data } = await apiClient.post<NewsletterMutationResponse>(
    API_ROUTES.newsletters,
    input,
  )
  return data
}

export async function updateNewsletterApiEntry(
  id: string,
  patch: Partial<NewsletterApiCreateInput>,
): Promise<NewsletterMutationResponse> {
  const { data } = await apiClient.put<NewsletterMutationResponse>(
    API_ROUTES.newsletterById(id),
    patch,
  )
  return data
}

export async function deleteNewsletterApiEntry(id: string): Promise<void> {
  await apiClient.delete(API_ROUTES.newsletterById(id))
}

export async function addNewsletterMedia(
  entryId: string,
  files: File[],
  metadata: NewsletterMediaUploadInput[],
): Promise<AddNewsletterMediaResponse> {
  files.forEach((file) => {
    assertValidResourceUploadFile(file)
  })

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

  const { data } = await apiClient.post<AddNewsletterMediaResponse>(
    `${API_ROUTES.newsletterById(entryId)}/media`,
    body,
  )

  return data
}

export async function updateNewsletterMedia(
  entryId: string,
  mediaId: string,
  input: UpdateNewsletterMediaInput,
): Promise<NewsletterMedia> {
  const { data } = await apiClient.patch<UpdateNewsletterMediaResponse>(
    `${API_ROUTES.newsletterById(entryId)}/media/${mediaId}`,
    input,
  )

  return newsletterApiMediaToLocal(data.media)
}

export async function getNewsletterMediaContent(
  entryId: string,
  mediaId: string,
): Promise<Blob> {
  const response = await apiClient.get<Blob>(
    API_ROUTES.newsletterMediaById(entryId, mediaId),
    { responseType: 'blob' },
  )

  return response.data
}

export async function reorderNewsletterMedia(
  entryId: string,
  mediaIds: number[],
): Promise<ReorderNewsletterMediaResponse> {
  const { data } = await apiClient.put<ReorderNewsletterMediaResponse>(
    `${API_ROUTES.newsletterById(entryId)}/media/order`,
    { media_ids: mediaIds } satisfies ReorderNewsletterMediaRequest,
  )

  return data
}

export async function deleteNewsletterMedia(
  entryId: string,
  mediaIds: number[],
): Promise<DeleteNewsletterMediaResponse> {
  const { data } = await apiClient.delete<DeleteNewsletterMediaResponse>(
    `${API_ROUTES.newsletterById(entryId)}/media`,
    { data: { media_ids: mediaIds } satisfies DeleteNewsletterMediaRequest },
  )

  return data
}
