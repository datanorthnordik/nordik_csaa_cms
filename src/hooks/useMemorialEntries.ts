import { useCallback, useState } from 'react'
import { memorialApi, type MemorialApiMutationInput } from '../api/memorialApi'
import type {
  MemorialEntrySummary,
  MemorialListFilters,
  MemorialListPageMeta,
} from '../lib/memorialTypes'

type UseMemorialEntriesState = {
  items: MemorialEntrySummary[]
  pagination: MemorialListPageMeta
  loading: boolean
  error: string | null
}

const defaultPagination: MemorialListPageMeta = {
  page: 1,
  pageSize: 10,
  totalItems: 0,
  totalPages: 0,
  hasNext: false,
  hasPrev: false,
}

export function useMemorialEntries(initialPageSize = 10) {
  const [state, setState] = useState<UseMemorialEntriesState>({
    items: [],
    pagination: { ...defaultPagination, pageSize: initialPageSize },
    loading: false,
    error: null,
  })

  const fetch = useCallback(async (filters: MemorialListFilters) => {
    setState((current) => ({ ...current, loading: true, error: null }))

    try {
      const response = await memorialApi.listMemorials(filters)
      setState({
        items: response.items,
        pagination: response.pagination,
        loading: false,
        error: null,
      })
      return response
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Failed to fetch memorial entries'
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
      input: MemorialApiMutationInput,
      portraitFile?: File | null,
      galleryFiles: File[] = [],
    ) => {
      const created = await memorialApi.createMemorial(input, portraitFile, galleryFiles)
      return memorialApi.getMemorial(String(created.id))
    },
    [],
  )

  const update = useCallback(
    async (
      id: string,
      input: MemorialApiMutationInput,
      portraitFile?: File | null,
      galleryFiles: File[] = [],
    ) => {
      await memorialApi.updateMemorial(id, input, portraitFile, galleryFiles)
      return memorialApi.getMemorial(id)
    },
    [],
  )

  const remove = useCallback(async (id: string) => {
    await memorialApi.deleteMemorial(id)
    setState((current) => ({
      ...current,
      items: current.items.filter((item) => item.id !== id),
      pagination: {
        ...current.pagination,
        totalItems: Math.max(0, current.pagination.totalItems - 1),
      },
    }))
  }, [])

  return {
    ...state,
    fetch,
    create,
    update,
    remove,
  }
}
