import type { PublishStatus } from '../components/cms/StatusBadge'
import type { PublishVisibility } from '../components/cms/PublishingControls'

export type PressStatus = PublishStatus
export type PressVisibility = PublishVisibility

export type PressMedia = {
  id: string
  fileName: string
  fileUrl?: string
  mimeType?: string
  fileSize?: number
}

export type PressEntry = {
  id: string
  title: string
  releaseDate: string
  categoryId: string | null
  sourceUrl: string
  contentHtml: string
  status: PressStatus
  visibility: PressVisibility
  publishAt: string | null
  coverImageUrl?: string
  media: PressMedia[]
  createdAt: string
  updatedAt: string
}

export type PressCategory = {
  id: string
  name: string
  createdAt: string
}

export type PressFormState = {
  title: string
  releaseDate: string
  categoryId: string
  sourceUrl: string
  contentHtml: string
  visibility: PressVisibility
  publishAt: string
  media: PressMedia[]
}

export type PressFormErrors = Partial<
  Record<keyof PressFormState | 'coverImage' | 'media', string>
>

export type PressListFilters = {
  searchTerm: string
  status: 'all' | PressStatus
  sortBy: 'releaseDate' | 'title' | 'updatedAt'
  sortOrder: 'asc' | 'desc'
}

export const defaultPressListFilters: PressListFilters = {
  searchTerm: '',
  status: 'all',
  sortBy: 'releaseDate',
  sortOrder: 'desc',
}
