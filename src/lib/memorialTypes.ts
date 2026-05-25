export type MemorialStatus = 'published' | 'draft' | 'review'

export type MemorialCategory = 'alumnus' | 'veteran' | 'founder' | 'friend'

export type MemorialPortraitAsset = {
  fileName: string
  mimeType: string
  fileSize: number
}

export type MemorialGalleryImage = {
  id: string
  fileName: string
  mimeType: string
  fileSize: number
}

export type MemorialEntrySummary = {
  id: string
  fullName: string
  affiliation: string
  category: MemorialCategory
  categoryLabel: string
  status: MemorialStatus
  dateOfBirth: string
  dateOfPassing: string
  createdAt: string
  updatedAt: string
  publishedAt?: string
}

export type MemorialEntry = MemorialEntrySummary & {
  biography: string
  portrait?: MemorialPortraitAsset
  galleryImages: MemorialGalleryImage[]
}

export type MemorialFormState = {
  fullName: string
  biography: string
  category: MemorialCategory | ''
  affiliation: string
  status: MemorialStatus
  dateOfBirth: string
  dateOfPassing: string
}

export type MemorialFormErrors = Partial<Record<keyof MemorialFormState, string>>

export type MemorialListPageMeta = {
  page: number
  pageSize: number
  totalItems: number
  totalPages: number
  hasNext: boolean
  hasPrev: boolean
}

export type MemorialListFilters = {
  searchTerm: string
  status: MemorialStatus | 'all'
  category: MemorialCategory | ''
  page: number
  pageSize: number
}

export const defaultMemorialListFilters: MemorialListFilters = {
  searchTerm: '',
  status: 'all',
  category: '',
  page: 1,
  pageSize: 10,
}

export const MEMORIAL_CATEGORIES: MemorialCategory[] = [
  'alumnus',
  'veteran',
  'founder',
  'friend',
]
