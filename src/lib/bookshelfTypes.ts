export type BookshelfListPageMeta = {
  page: number
  pageSize: number
  totalItems: number
  totalPages: number
  hasNext: boolean
  hasPrev: boolean
}

export type BookshelfListFilters = {
  page: number
  pageSize: number
  searchTerm: string
}

export type BookshelfListSummary = {
  withCoverCount: number
  withoutCoverCount: number
}

export type BookshelfFormState = {
  author: string
  title: string
  bookLink: string
  authorBio: string
  bookTeaser: string
  description: string
}

export type BookshelfFormErrors = Partial<
  Record<keyof BookshelfFormState | 'bookFile' | 'authorImage' | 'coverImage', string>
>

export type SaveBookshelfInput = BookshelfFormState & {
  removeAuthorImage?: boolean
  removeCoverImage?: boolean
}

export type BookshelfEntry = {
  id: string
  author: string
  title: string
  bookLink: string
  authorBio: string
  bookTeaser: string
  description: string
  bookFileName: string
  bookMimeType: string
  bookFileSize: number
  bookContentUrl: string
  authorImageFileName: string
  authorImageMimeType: string
  authorImageFileSize: number
  hasAuthorImage: boolean
  authorImageContentUrl: string
  coverImageFileName: string
  coverImageMimeType: string
  coverImageFileSize: number
  hasCoverImage: boolean
  coverImageContentUrl: string
  createdAt: string
  updatedAt: string
}

export const defaultBookshelfListFilters: BookshelfListFilters = {
  page: 1,
  pageSize: 10,
  searchTerm: '',
}
