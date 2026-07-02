import { API_ROUTES } from '../constants/api'
import {
  assertValidBookshelfAuthorImageFile,
  assertValidBookshelfBookFile,
  assertValidBookshelfCoverFile,
} from '../lib/bookshelfUpload'
import type {
  BookshelfEntry,
  BookshelfListFilters,
  BookshelfListPageMeta,
  BookshelfListSummary,
} from '../lib/bookshelfTypes'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

type BookshelfApiEntry = {
  id: number
  author: string
  title: string
  book_link: string
  author_bio: string
  book_teaser: string
  description: string
  book_file_name: string
  book_mime_type: string
  book_file_size: number
  book_content_url: string
  author_image_file_name: string
  author_image_mime_type: string
  author_image_file_size: number
  has_author_image: boolean
  author_image_content_url: string
  cover_image_file_name: string
  cover_image_mime_type: string
  cover_image_file_size: number
  has_cover_image: boolean
  cover_image_content_url: string
  created_at: string
  updated_at: string
}

type BookshelfApiListResponse = {
  items: BookshelfApiEntry[]
  pagination: {
    page: number
    page_size: number
    total_items: number
    total_pages: number
    has_next: boolean
    has_prev: boolean
  }
  summary: {
    with_cover_count: number
    without_cover_count: number
  }
  applied_filters: {
    page: number
    page_size: number
    search_term: string
  }
}

type BookshelfApiMutationInput = {
  author: string
  title: string
  bookLink: string
  authorBio: string
  bookTeaser: string
  description: string
  removeAuthorImage?: boolean
  removeCoverImage?: boolean
}

type BookshelfApiMutationResponse = {
  message: string
  book: {
    id: number
    author: string
    title: string
    updated_at: string
  }
}

function buildListQuery(filters: BookshelfListFilters) {
  const params = new URLSearchParams()

  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }

  return params
}

function buildBookshelfMutationBody(
  input: BookshelfApiMutationInput,
  bookFile?: File,
  authorImageFile?: File,
  coverFile?: File,
) {
  const payload = {
    author: input.author,
    title: input.title,
    book_link: input.bookLink,
    author_bio: input.authorBio,
    book_teaser: input.bookTeaser,
    description: input.description,
    remove_author_image: Boolean(input.removeAuthorImage),
    remove_cover_image: Boolean(input.removeCoverImage),
    book_upload: bookFile
      ? {
          file_name: bookFile.name,
          mime_type: bookFile.type || 'application/octet-stream',
          file_size: bookFile.size,
        }
      : undefined,
    author_image: authorImageFile
      ? {
          file_name: authorImageFile.name,
          mime_type: authorImageFile.type || 'application/octet-stream',
          file_size: authorImageFile.size,
        }
      : undefined,
    cover_image: coverFile
      ? {
          file_name: coverFile.name,
          mime_type: coverFile.type || 'application/octet-stream',
          file_size: coverFile.size,
        }
      : undefined,
  }

  if (!bookFile && !authorImageFile && !coverFile) {
    return payload
  }

  if (bookFile) {
    assertValidBookshelfBookFile(bookFile)
  }
  if (authorImageFile) {
    assertValidBookshelfAuthorImageFile(authorImageFile)
  }
  if (coverFile) {
    assertValidBookshelfCoverFile(coverFile)
  }

  const files = []

  if (bookFile) {
    files.push({
      fieldName: 'book_file',
      file: bookFile,
      fileName: bookFile.name,
    })
  }
  if (authorImageFile) {
    files.push({
      fieldName: 'author_image_file',
      file: authorImageFile,
      fileName: authorImageFile.name,
    })
  }
  if (coverFile) {
    files.push({
      fieldName: 'cover_image_file',
      file: coverFile,
      fileName: coverFile.name,
    })
  }

  return buildMultipartPayload(payload, files)
}

function mapBookshelfEntry(entry: BookshelfApiEntry): BookshelfEntry {
  return {
    id: String(entry.id),
    author: entry.author,
    title: entry.title,
    bookLink: entry.book_link,
    authorBio: entry.author_bio,
    bookTeaser: entry.book_teaser,
    description: entry.description,
    bookFileName: entry.book_file_name,
    bookMimeType: entry.book_mime_type,
    bookFileSize: entry.book_file_size,
    bookContentUrl: entry.book_content_url,
    authorImageFileName: entry.author_image_file_name,
    authorImageMimeType: entry.author_image_mime_type,
    authorImageFileSize: entry.author_image_file_size,
    hasAuthorImage: entry.has_author_image,
    authorImageContentUrl: entry.author_image_content_url,
    coverImageFileName: entry.cover_image_file_name,
    coverImageMimeType: entry.cover_image_mime_type,
    coverImageFileSize: entry.cover_image_file_size,
    hasCoverImage: entry.has_cover_image,
    coverImageContentUrl: entry.cover_image_content_url,
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

function mapPagination(
  pagination: BookshelfApiListResponse['pagination'],
): BookshelfListPageMeta {
  return {
    page: pagination.page,
    pageSize: pagination.page_size,
    totalItems: pagination.total_items,
    totalPages: pagination.total_pages,
    hasNext: pagination.has_next,
    hasPrev: pagination.has_prev,
  }
}

function mapSummary(summary: BookshelfApiListResponse['summary']): BookshelfListSummary {
  return {
    withCoverCount: summary.with_cover_count,
    withoutCoverCount: summary.without_cover_count,
  }
}

export const bookshelfApi = {
  async listBooks(filters: BookshelfListFilters) {
    const response = await apiClient.get<BookshelfApiListResponse>(API_ROUTES.bookshelf, {
      params: buildListQuery(filters),
    })

    return {
      items: response.data.items.map(mapBookshelfEntry),
      pagination: mapPagination(response.data.pagination),
      summary: mapSummary(response.data.summary),
      appliedFilters: {
        page: response.data.applied_filters.page,
        pageSize: response.data.applied_filters.page_size,
        searchTerm: response.data.applied_filters.search_term,
      },
    }
  },

  async getBook(id: string) {
    const response = await apiClient.get<BookshelfApiEntry>(API_ROUTES.bookshelfById(id))
    return mapBookshelfEntry(response.data)
  },

  async getBookContent(id: string) {
    const response = await apiClient.get<Blob>(API_ROUTES.bookshelfBookContentById(id), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async getAuthorImageContent(id: string) {
    const response = await apiClient.get<Blob>(API_ROUTES.bookshelfAuthorImageContentById(id), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async getCoverImageContent(id: string) {
    const response = await apiClient.get<Blob>(API_ROUTES.bookshelfCoverContentById(id), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createBook(
    input: BookshelfApiMutationInput,
    bookFile?: File,
    authorImageFile?: File,
    coverFile?: File,
  ) {
    const response = await apiClient.post<BookshelfApiMutationResponse>(
      API_ROUTES.bookshelf,
      buildBookshelfMutationBody(input, bookFile, authorImageFile, coverFile),
    )
    return response.data.book
  },

  async updateBook(
    id: string,
    input: BookshelfApiMutationInput,
    bookFile?: File,
    authorImageFile?: File,
    coverFile?: File,
  ) {
    const response = await apiClient.put<BookshelfApiMutationResponse>(
      API_ROUTES.bookshelfById(id),
      buildBookshelfMutationBody(input, bookFile, authorImageFile, coverFile),
    )
    return response.data.book
  },

  async deleteBook(id: string) {
    await apiClient.delete(API_ROUTES.bookshelfById(id))
  },
}
