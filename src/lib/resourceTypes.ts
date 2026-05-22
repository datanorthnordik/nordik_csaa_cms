export type ResourceCategory =
  | 'brand_identity'
  | 'governance_legal'
  | 'training_manuals'
  | 'media_kits'

export type ResourceVisibility = 'public' | 'internal'

export type ResourceFileType =
  | 'all'
  | 'pdf'
  | 'document'
  | 'presentation'
  | 'spreadsheet'
  | 'image'
  | 'vector'
  | 'other'

export type ResourceCategoryOption = {
  id: ResourceCategory
  label: string
}

export type ResourceCategoryCount = {
  category: ResourceCategory
  label: string
  count: number
}

export type ResourceEntry = {
  id: string
  name: string
  category: ResourceCategory
  categoryLabel: string
  visibility: ResourceVisibility
  fileName: string
  mimeType: string
  fileSize: number
  contentUrl: string
  createdAt: string
  updatedAt: string
}

export type ResourceListPageMeta = {
  page: number
  pageSize: number
  totalItems: number
  totalPages: number
  hasNext: boolean
  hasPrev: boolean
}

export type ResourceListFilters = {
  page: number
  pageSize: number
  searchTerm: string
  category: ResourceCategory | ''
  fileType: ResourceFileType
}

export type ResourceFormState = {
  name: string
  category: ResourceCategory | ''
  visibility: ResourceVisibility
}

export type ResourceFormErrors = Partial<Record<keyof ResourceFormState | 'file', string>>

export const resourceCategoryOptions: ResourceCategoryOption[] = [
  { id: 'brand_identity', label: 'Brand Identity' },
  { id: 'governance_legal', label: 'Governance & Legal' },
  { id: 'training_manuals', label: 'Training & Manuals' },
  { id: 'media_kits', label: 'Media Kits' },
]

export const resourceFileTypeOptions: Array<{ value: ResourceFileType; label: string }> = [
  { value: 'all', label: 'All Types' },
  { value: 'pdf', label: 'PDF' },
  { value: 'document', label: 'Document' },
  { value: 'presentation', label: 'Presentation' },
  { value: 'spreadsheet', label: 'Spreadsheet' },
  { value: 'image', label: 'Image' },
  { value: 'vector', label: 'Vector' },
  { value: 'other', label: 'Other' },
]

export const defaultResourceListFilters: ResourceListFilters = {
  page: 1,
  pageSize: 6,
  searchTerm: '',
  category: '',
  fileType: 'all',
}
