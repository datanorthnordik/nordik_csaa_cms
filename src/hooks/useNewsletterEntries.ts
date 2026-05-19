import { useCallback, useState } from 'react'
import {
  createNewsletterApiEntry,
  deleteNewsletterApiEntry,
  fetchNewsletterEntries,
  fetchNewsletterEntry,
  newsletterApiEntryToLocal,
  updateNewsletterApiEntry,
} from '../api/newslettersApi'
import type {
  NewsletterEntry,
  NewsletterListFilters,
  NewsletterStatus,
  NewsletterVisibility,
} from '../lib/newsletterTypes'

type NewsletterPersistedInput = Omit<NewsletterEntry, 'id' | 'createdAt' | 'updatedAt' | 'media'>
type CreateInput = NewsletterPersistedInput
type UpdateInput = Partial<NewsletterPersistedInput>

export interface UseNewsletterEntriesState {
  entries: NewsletterEntry[]
  loading: boolean
  error: string | null
  total: number
  page: number
  pageSize: number
  totalPages: number
}

function mapSortBy(sortBy: NewsletterListFilters['sortBy']) {
  switch (sortBy) {
    case 'title':
      return 'title'
    case 'updatedAt':
      return 'updated_at'
    default:
      return 'send_date'
  }
}

export function useNewsletterEntries(page: number = 1, pageSize: number = 20) {
  const [state, setState] = useState<UseNewsletterEntriesState>({
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
      filters?: {
        status?: NewsletterStatus
        visibility?: NewsletterVisibility
        search?: string
        sortBy?: NewsletterListFilters['sortBy']
        sortOrder?: NewsletterListFilters['sortOrder']
      },
    ) => {
      setState((prev) => ({ ...prev, loading: true, error: null }))

      try {
        const params: Record<string, string | number> = {
          page: currentPage,
          page_size: currentPageSize,
          sort_by: mapSortBy(filters?.sortBy ?? 'sendDate'),
          sort_order: filters?.sortOrder ?? 'desc',
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

        const response = await fetchNewsletterEntries(params)
        const entries = response.items.map(newsletterApiEntryToLocal)

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
        const message =
          err instanceof Error ? err.message : 'Failed to fetch newsletter entries'
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

  const create = useCallback(async (input: CreateInput) => {
    try {
      const result = await createNewsletterApiEntry({
        title: input.title,
        category: input.category,
        send_date: input.sendDate,
        content_html: input.contentHtml,
        status: input.status,
        visibility: input.visibility,
        publish_at: input.publishAt,
      })

      const created = await fetchNewsletterEntry(String(result.entry.id))
      setState((prev) => ({
        ...prev,
        entries: [created, ...prev.entries].slice(0, prev.pageSize),
        total: prev.total + 1,
      }))

      return created
    } catch (err) {
      const message =
        err instanceof Error ? err.message : 'Failed to create newsletter entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const update = useCallback(async (id: string, patch: UpdateInput) => {
    try {
      await updateNewsletterApiEntry(id, {
        title: patch.title,
        category: patch.category,
        send_date: patch.sendDate,
        content_html: patch.contentHtml,
        status: patch.status,
        visibility: patch.visibility,
        publish_at: patch.publishAt,
      })

      const updated = await fetchNewsletterEntry(id)
      setState((prev) => ({
        ...prev,
        entries: prev.entries.map((entry) => (entry.id === id ? updated : entry)),
      }))

      return updated
    } catch (err) {
      const message =
        err instanceof Error ? err.message : 'Failed to update newsletter entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const remove = useCallback(async (id: string) => {
    try {
      await deleteNewsletterApiEntry(id)
      setState((prev) => ({
        ...prev,
        entries: prev.entries.filter((entry) => entry.id !== id),
        total: prev.total - 1,
      }))
    } catch (err) {
      const message =
        err instanceof Error ? err.message : 'Failed to delete newsletter entry'
      setState((prev) => ({ ...prev, error: message }))
      throw err
    }
  }, [])

  const get = useCallback(
    (id: string) => state.entries.find((entry) => entry.id === id),
    [state.entries],
  )

  return {
    ...state,
    fetch,
    create,
    update,
    remove,
    get,
  }
}
