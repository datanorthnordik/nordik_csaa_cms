import { API_ROUTES } from '../constants/api'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type EventStatus = 'published' | 'draft'
export type EventDateRange =
  | 'custom'
  | 'next_30_days'
  | 'last_30_days'
  | 'today'
  | 'this_month'
  | 'upcoming'
export type EventSortBy =
  | 'title'
  | 'start_at'
  | 'created_at'
  | 'updated_at'
  | 'published'
export type EventSortOrder = 'asc' | 'desc'
export type EventType =
  | 'single_day_all_day'
  | 'single_day_partial'
  | 'multi_day_all_day'
  | 'multi_day_partial'
export type LocationMode = 'none' | 'to_be_determined' | 'address'
export type RecurrenceType = 'scheduled' | 'recurring'
export type RecurrenceFrequency = 'daily' | 'weekly' | 'monthly' | 'yearly'

export type EventAddress = {
  id: number
  name: string
  address_line_1: string
  address_line_2: string
  city: string
  province_state: string
  postal_code: string
  country: string
  is_saved: boolean
  created_at: string
  updated_at: string
}

export type Gallery = {
  id: number
  name: string
  created_at: string
  updated_at: string
}

export type EventMedia = {
  id: number
  event_id: number
  media_role: 'display_image' | 'attachment'
  display_name: string
  gcp_object_key: string
  file_url: string
  fetch_url?: string
  storage_uri?: string
  mime_type: string
  file_size: number
  sort_order: number
  created_at: string
  updated_at: string
}

export type EventOccurrence = {
  id: number
  event_id: number
  occurrence_start_at: string
  occurrence_end_at?: string | null
  occurrence_kind: string
  created_at: string
  updated_at: string
}

export type EventListItem = {
  id: number
  title: string
  categories: string[]
  status: EventStatus
  published: boolean
  event_type: EventType
  start_at: string
  end_at?: string | null
  date_display?: string
  created_at: string
  updated_at: string
}

export type EventListPageMeta = {
  page: number
  page_size: number
  total_items: number
  total_pages: number
  has_next: boolean
  has_prev: boolean
}

export type EventListFilters = {
  page: number
  pageSize: number
  searchTerm: string
  statuses: EventStatus[]
  startDate: string
  endDate: string
  dateRange: EventDateRange
  sortBy: EventSortBy
  sortOrder: EventSortOrder
}

export type EventListResponse = {
  items: EventListItem[]
  pagination: EventListPageMeta
  applied_filters: {
    page: number
    page_size: number
    search_term: string
    statuses: EventStatus[]
    start_date?: string | null
    end_date?: string | null
    date_range: EventDateRange
    sort_by: EventSortBy
    sort_order: EventSortOrder
  }
}

export type EventDetailResponse = {
  id: number
  title: string
  show_title: boolean
  categories: string[]
  event_type: EventType
  start_at: string
  end_at?: string | null
  date_display?: string
  privacy_type: 'public' | 'private'
  private_audiences: string[]
  published: boolean
  request_review: boolean
  review_email_list: string[]
  teaser: string
  description_html: string
  contact_name: string
  contact_email: string
  contact_phone: string
  contact_ext: string
  contact_fax: string
  location_mode: LocationMode
  address?: EventAddress | null
  show_display_image_when_viewing: boolean
  gallery_id?: number | null
  registration_enabled: boolean
  registration_start_at?: string | null
  registration_end_at?: string | null
  registration_url: string
  repeat_enabled: boolean
  recurrence_type?: RecurrenceType | '' | null
  recurrence_frequency?: RecurrenceFrequency | '' | null
  recurrence_interval: number
  recurrence_until?: string | null
  recurrence_rule?: unknown
  occurrences: EventOccurrence[]
  display_image?: EventMedia | null
  attachments: EventMedia[]
  created_by?: number | null
  created_at: string
  updated_at: string
}

export type EventUploadInput = {
  display_name?: string
  file_name?: string
  mime_type?: string
  file_url?: string
  object_key?: string
  gcp_object_key?: string
  storage_uri?: string
}

export type EventAddressInput = {
  id?: number
  name: string
  address_line_1: string
  address_line_2: string
  city: string
  province_state: string
  postal_code: string
  country: string
  is_saved: boolean
}

export type EventOccurrenceInput = {
  occurrence_start_at: string
  occurrence_end_at?: string | null
  occurrence_kind: string
}

export type SaveEventPayload = {
  title: string
  show_title: boolean
  categories: string[]
  event_type: EventType
  start_at: string
  end_at?: string | null
  privacy_type: 'public' | 'private'
  private_audiences: string[]
  published: boolean
  request_review: boolean
  review_email_list: string[]
  teaser: string
  description_html: string
  contact_name: string
  contact_email: string
  contact_phone: string
  contact_ext: string
  contact_fax: string
  location_mode: LocationMode
  address?: EventAddressInput
  show_display_image_when_viewing: boolean
  gallery_id?: number | null
  registration_enabled: boolean
  registration_start_at?: string | null
  registration_end_at?: string | null
  registration_url: string
  repeat_enabled: boolean
  recurrence_type?: RecurrenceType | ''
  recurrence_frequency?: RecurrenceFrequency | ''
  recurrence_interval: number
  recurrence_until?: string | null
  recurrence_rule?: unknown
  occurrences: EventOccurrenceInput[]
  display_image?: EventUploadInput
  attachments: EventUploadInput[]
}

export type SaveEventRequest = SaveEventPayload & {
  displayImageFile?: File
  attachmentFiles?: Array<File | null | undefined>
}

export type EventMutationResponse = {
  message: string
  event: {
    id: number
    title: string
    published: boolean
  }
}

function buildListQuery(filters: EventListFilters) {
  const params = new URLSearchParams()
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
  params.set('date_range', filters.dateRange)
  params.set('sort_by', filters.sortBy)
  params.set('sort_order', filters.sortOrder)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }
  for (const status of filters.statuses) {
    params.append('status', status)
  }
  if (filters.startDate) {
    params.set('start_date', filters.startDate)
  }
  if (filters.endDate) {
    params.set('end_date', filters.endDate)
  }

  return params
}

export const eventsApi = {
  async listEvents(filters: EventListFilters) {
    const response = await apiClient.get<EventListResponse>(API_ROUTES.events, {
      params: buildListQuery(filters),
    })
    return response.data
  },

  async getEvent(id: number) {
    const response = await apiClient.get<EventDetailResponse>(API_ROUTES.eventById(id))
    return response.data
  },

  async fetchEventMediaContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async getSavedLocations() {
    const response = await apiClient.get<{ items: EventAddress[] }>(
      API_ROUTES.eventLocations,
    )
    return response.data.items
  },

  async getGalleries() {
    const response = await apiClient.get<{ items: Gallery[] }>(
      API_ROUTES.eventGalleries,
    )
    return response.data.items
  },

  async createEvent(request: SaveEventRequest) {
    const { displayImageFile, attachmentFiles, ...payload } = request
    const hasFiles = Boolean(displayImageFile) || Boolean(attachmentFiles?.some(Boolean))
    const body = hasFiles
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'display_image_file',
            file: displayImageFile,
            fileName: displayImageFile?.name,
          },
          ...(attachmentFiles ?? []).map((file, index) => ({
            fieldName: `attachments[${index}].file`,
            file,
            fileName: file?.name,
          })),
        ])
      : payload
    const response = await apiClient.post<EventMutationResponse>(
      API_ROUTES.events,
      body,
    )
    return response.data
  },

  async updateEvent(id: number, request: SaveEventRequest) {
    const { displayImageFile, attachmentFiles, ...payload } = request
    const hasFiles = Boolean(displayImageFile) || Boolean(attachmentFiles?.some(Boolean))
    const body = hasFiles
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'display_image_file',
            file: displayImageFile,
            fileName: displayImageFile?.name,
          },
          ...(attachmentFiles ?? []).map((file, index) => ({
            fieldName: `attachments[${index}].file`,
            file,
            fileName: file?.name,
          })),
        ])
      : payload
    const response = await apiClient.put<EventMutationResponse>(
      API_ROUTES.eventById(id),
      body,
    )
    return response.data
  },

  async deleteEvent(id: number) {
    await apiClient.delete(API_ROUTES.eventById(id))
  },

  async deleteEventDocument(eventId: number, mediaId: number) {
    await apiClient.delete(API_ROUTES.eventDocumentById(eventId, mediaId))
  },

  async deleteEventPhoto(eventId: number, mediaId: number) {
    await apiClient.delete(API_ROUTES.eventPhotoById(eventId, mediaId))
  },
}
