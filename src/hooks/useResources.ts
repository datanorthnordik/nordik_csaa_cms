import { useCallback, useState } from 'react'
import { resourcesApi } from '../api/resourcesApi'
import type {
  ResourceCategoryCount,
  ResourceEntry,
  ResourceListFilters,
  ResourceListPageMeta,
  ResourceVisibility,
} from '../lib/resourceTypes'

type PersistedResourceInput = Pick<
  ResourceEntry,
  'name' | 'description' | 'category' | 'visibility' | 'linkUrl'
>

type UseResourcesState = {
  items: ResourceEntry[]
  pagination: ResourceListPageMeta
  categoryCounts: ResourceCategoryCount[]
  loading: boolean
  error: string | null
}

const defaultPagination: ResourceListPageMeta = {
  page: 1,
  pageSize: 6,
  totalItems: 0,
  totalPages: 0,
  hasNext: false,
  hasPrev: false,
}

export function useResources(initialPageSize = 6) {
  const [state, setState] = useState<UseResourcesState>({
    items: [],
    pagination: { ...defaultPagination, pageSize: initialPageSize },
    categoryCounts: [],
    loading: false,
    error: null,
  })

  const fetch = useCallback(async (filters: ResourceListFilters) => {
    setState((current) => ({ ...current, loading: true, error: null }))

    try {
      const response = await resourcesApi.listResources(filters)

      setState({
        items: response.items,
        pagination: response.pagination,
        categoryCounts: response.summary.categoryCounts,
        loading: false,
        error: null,
      })

      return response
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Failed to fetch resources'

      setState((current) => ({
        ...current,
        loading: false,
        error: message,
      }))

      throw error
    }
  }, [])

  const create = useCallback(
    async (input: PersistedResourceInput, file?: File) => {
      const created = await resourcesApi.createResource(
        {
          name: input.name,
          description: input.description,
          category: input.category,
          visibility: input.visibility as ResourceVisibility,
          linkUrl: input.linkUrl,
        },
        file,
      )

      return resourcesApi.getResource(String(created.id))
    },
    [],
  )

  const update = useCallback(
    async (id: string, input: PersistedResourceInput, file?: File) => {
      await resourcesApi.updateResource(
        id,
        {
          name: input.name,
          description: input.description,
          category: input.category,
          visibility: input.visibility as ResourceVisibility,
          linkUrl: input.linkUrl,
        },
        file,
      )

      return resourcesApi.getResource(id)
    },
    [],
  )

  const remove = useCallback(async (id: string) => {
    await resourcesApi.deleteResource(id)

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