import { useCallback, useState } from 'react'
import { bookshelfApi } from '../api/bookshelfApi'
import type {
  BookshelfEntry,
  BookshelfListFilters,
  BookshelfListPageMeta,
  BookshelfListSummary,
  SaveBookshelfInput,
} from '../lib/bookshelfTypes'

type UseBookshelfState = {
  items: BookshelfEntry[]
  pagination: BookshelfListPageMeta
  summary: BookshelfListSummary
  loading: boolean
  error: string | null
}

const defaultPagination: BookshelfListPageMeta = {
  page: 1,
  pageSize: 10,
  totalItems: 0,
  totalPages: 0,
  hasNext: false,
  hasPrev: false,
}

const defaultSummary: BookshelfListSummary = {
  withCoverCount: 0,
  withoutCoverCount: 0,
}

export function useBookshelf(initialPageSize = 10) {
  const [state, setState] = useState<UseBookshelfState>({
    items: [],
    pagination: { ...defaultPagination, pageSize: initialPageSize },
    summary: defaultSummary,
    loading: false,
    error: null,
  })

  const fetch = useCallback(async (filters: BookshelfListFilters) => {
    setState((current) => ({ ...current, loading: true, error: null }))

    try {
      const response = await bookshelfApi.listBooks(filters)

      setState({
        items: response.items,
        pagination: response.pagination,
        summary: response.summary,
        loading: false,
        error: null,
      })

      return response
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to fetch bookshelf entries'

      setState((current) => ({
        ...current,
        loading: false,
        error: message,
      }))

      throw error
    }
  }, [])

  const create = useCallback(
    async (
      input: SaveBookshelfInput,
      bookFile?: File,
      authorImageFile?: File,
      coverFile?: File,
    ) => {
      const created = await bookshelfApi.createBook(
        {
          author: input.author,
          title: input.title,
          bookLink: input.bookLink,
          authorBio: input.authorBio,
          bookTeaser: input.bookTeaser,
          description: input.description,
          removeAuthorImage: input.removeAuthorImage,
          removeCoverImage: input.removeCoverImage,
        },
        bookFile,
        authorImageFile,
        coverFile,
      )

      return bookshelfApi.getBook(String(created.id))
    },
    [],
  )

  const update = useCallback(
    async (
      id: string,
      input: SaveBookshelfInput,
      bookFile?: File,
      authorImageFile?: File,
      coverFile?: File,
    ) => {
      await bookshelfApi.updateBook(
        id,
        {
          author: input.author,
          title: input.title,
          bookLink: input.bookLink,
          authorBio: input.authorBio,
          bookTeaser: input.bookTeaser,
          description: input.description,
          removeAuthorImage: input.removeAuthorImage,
          removeCoverImage: input.removeCoverImage,
        },
        bookFile,
        authorImageFile,
        coverFile,
      )

      return bookshelfApi.getBook(id)
    },
    [],
  )

  const remove = useCallback(async (id: string) => {
    await bookshelfApi.deleteBook(id)

    setState((current) => {
      const removedItem = current.items.find((item) => item.id === id)

      return {
        ...current,
        items: current.items.filter((item) => item.id !== id),
        pagination: {
          ...current.pagination,
          totalItems: Math.max(0, current.pagination.totalItems - 1),
        },
        summary: removedItem
          ? {
              withCoverCount: Math.max(
                0,
                current.summary.withCoverCount - (removedItem.hasCoverImage ? 1 : 0),
              ),
              withoutCoverCount: Math.max(
                0,
                current.summary.withoutCoverCount - (removedItem.hasCoverImage ? 0 : 1),
              ),
            }
          : current.summary,
      }
    })
  }, [])

  return {
    ...state,
    fetch,
    create,
    update,
    remove,
  }
}
