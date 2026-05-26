

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


export type ResourceCategory = 'educational' | 'media' | 'link' | 'report'

export const resourceCategoryOptions: Array<{
  id: ResourceCategory
  label: string
}> = [
  { id: 'educational', label: 'Educational' },
  { id: 'media', label: 'Media' },
  { id: 'link', label: 'Link' },
  { id: 'report', label: 'Report' },
]

export type ResourceFormState = {
  name: string
  description: string
  category: ResourceCategory | ''
  visibility: 'public' | 'internal'
  linkUrl: string
}

export type ResourceFormErrors = Partial<Record<keyof ResourceFormState | 'file', string>>

export type SaveResourceInput = {
  name: string
  description: string
  category: ResourceCategory
  visibility: 'public' | 'internal'
  linkUrl?: string
}

export type ResourceEntry = {
  id: string
  name: string
  description: string
  category: ResourceCategory
  categoryLabel: string
  visibility: 'public' | 'internal'
  linkUrl: string
  fileName: string
  mimeType: string
  fileSize: number
  hasDocument: boolean
  contentUrl: string
  createdAt: string
  updatedAt: string
}

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
  pageSize: 10,
  searchTerm: '',
  category: '',
  fileType: 'all',
}
