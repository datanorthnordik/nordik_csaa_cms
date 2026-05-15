import { useCallback, useState } from 'react'
import {
  createPressApiEntry,
  deletePressApiEntry,
  fetchPressEntries,
  fetchPressEntry,
  pressApiEntryToLocal,
  updatePressApiEntry,
} from '../api/pressApi'
import type { PressEntry, PressStatus, PressVisibility } from '../lib/pressTypes'

type PressPersistedInput = Omit<PressEntry, 'id' | 'createdAt' | 'updatedAt' | 'media'>
type CreateInput = PressPersistedInput
type UpdateInput = Partial<PressPersistedInput>
type MutationOptions = {
  coverImageFile?: File
  removeCoverImage?: boolean
}

export interface UsePressEntriesState {
  entries: PressEntry[]
  loading: boolean
  error: string | null
  total: number
  page: number
  pageSize: number
  totalPages: number
}

export function usePressEntries(page: number = 1, pageSize: number = 20) {
  const [state, setState] = useState<UsePressEntriesState>({
    entries: [],
    loading: false,
    error: null,
    total: 0,
    page,
    pageSize,
    totalPages: 0,
  })

  const fetch = useCallback(
    async (
      currentPage: number = page,
      currentPageSize: number = pageSize,
      filters?: { status?: PressStatus; visibility?: PressVisibility; search?: string },
    ) => {
      setState((prev) => ({ ...prev, loading: true, error: null }))

      try {
        const params: Record<string, string | number> = {
          page: currentPage,
          page_size: currentPageSize,
        }

        if (filters?.status) {
          params.status = filters.status
        }
        if (filters?.visibility) {
          params.visibility = filters.visibility
        }
        if (filters?.search) {
          params.search = filters.search
        }

        const response = await fetchPressEntries(params)
        const entries = response.items.map(pressApiEntryToLocal)

        setState((prev) => ({
          ...prev,
          entries,
          total: response.total,
          page: response.page,
          pageSize: response.page_size,
          totalPages: response.total_pages,
          loading: false,
        }))
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Failed to fetch press entries'
        setState((prev) => ({
          ...prev,
          entries: [],
          loading: false,
          error: message,
        }))
        throw err
      }
    },
    [page, pageSize],
  )

  const create = useCallback(async (input: CreateInput, options: MutationOptions = {}) => {
    const { coverImageFile, removeCoverImage = false } = options
    try {
      const result = await createPressApiEntry(
        {
          title: input.title,
          release_date: input.releaseDate,
          category_id: input.categoryId ? Number.parseInt(input.categoryId, 10) : null,
          source_url: input.sourceUrl,
          content_html: input.contentHtml,
          status: input.status,
          visibility: input.visibility,
          publish_at: input.publishAt,
          remove_cover_image: removeCoverImage || undefined,
        },
        coverImageFile,
      )

      const created = await fetchPressEntry(String(result.entry.id))
      setState((prev) => ({
        ...prev,
        entries: [created, ...prev.entries].slice(0, prev.pageSize),
        total: prev.total + 1,
      }))

      return created
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to create press entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const update = useCallback(async (id: string, patch: UpdateInput, options: MutationOptions = {}) => {
    const { coverImageFile, removeCoverImage = false } = options
    try {
      await updatePressApiEntry(
        id,
        {
          title: patch.title,
          release_date: patch.releaseDate,
          category_id: patch.categoryId ? Number.parseInt(patch.categoryId, 10) : null,
          source_url: patch.sourceUrl,
          content_html: patch.contentHtml,
          status: patch.status,
          visibility: patch.visibility,
          publish_at: patch.publishAt,
          remove_cover_image: removeCoverImage || undefined,
        },
        coverImageFile,
      )

      const updated = await fetchPressEntry(id)
      setState((prev) => ({
        ...prev,
        entries: prev.entries.map((entry) => (entry.id === id ? updated : entry)),
      }))

      return updated
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to update press entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const remove = useCallback(async (id: string) => {
    try {
      await deletePressApiEntry(id)
      setState((prev) => ({
        ...prev,
        entries: prev.entries.filter((entry) => entry.id !== id),
        total: prev.total - 1,
      }))
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to delete press entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const get = useCallback((id: string) => {
    return state.entries.find((entry) => entry.id === id)
  }, [state.entries])

  return {
    ...state,
    fetch,
    create,
    update,
    remove,
    get,
  }
}
