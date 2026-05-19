import type { PublishStatus } from '../components/cms/StatusBadge'
import type { PublishVisibility } from '../components/cms/PublishingControls'

export type NewsletterStatus = PublishStatus
export type NewsletterVisibility = PublishVisibility
export type NewsletterCategory = 'csaa' | 'cst' | ''

export type NewsletterMedia = {
  id: string
  fileName: string
  fileUrl?: string
  mimeType?: string
  fileSize?: number
}

export type NewsletterEntry = {
  id: string
  title: string
  category: NewsletterCategory
  sendDate: string
  contentHtml: string
  status: NewsletterStatus
  visibility: NewsletterVisibility
  publishAt: string | null
  media: NewsletterMedia[]
  createdAt: string
  updatedAt: string
}

export type NewsletterFormState = {
  title: string
  category: NewsletterCategory
  sendDate: string
  contentHtml: string
  visibility: NewsletterVisibility
  publishAt: string
  media: NewsletterMedia[]
}

export type NewsletterFormErrors = Partial<Record<keyof NewsletterFormState, string>>

export type NewsletterListFilters = {
  searchTerm: string
  status: 'all' | NewsletterStatus
  sortBy: 'sendDate' | 'title' | 'updatedAt'
  sortOrder: 'asc' | 'desc'
}

export const defaultNewsletterListFilters: NewsletterListFilters = {
  searchTerm: '',
  status: 'all',
  sortBy: 'sendDate',
  sortOrder: 'desc',
}
