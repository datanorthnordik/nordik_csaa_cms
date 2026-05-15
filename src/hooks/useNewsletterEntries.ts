import { useCallback, useEffect, useState } from 'react'
import {
  createNewsletterEntry,
  deleteNewsletterEntry,
  loadNewsletterEntries,
  NEWSLETTER_ENTRIES_STORAGE_KEY,
  updateNewsletterEntry,
} from '../lib/newsletterEntriesStorage'
import type { NewsletterEntry } from '../lib/newsletterTypes'

type CreateInput = Omit<NewsletterEntry, 'id' | 'createdAt' | 'updatedAt'>
type UpdateInput = Partial<Omit<NewsletterEntry, 'id' | 'createdAt'>>

export function useNewsletterEntries() {
  const [entries, setEntries] = useState<NewsletterEntry[]>(() => loadNewsletterEntries())

  useEffect(() => {
    function handleStorage(event: StorageEvent) {
      if (event.key !== NEWSLETTER_ENTRIES_STORAGE_KEY) {
        return
      }
      setEntries(loadNewsletterEntries())
    }
    window.addEventListener('storage', handleStorage)
    return () => window.removeEventListener('storage', handleStorage)
  }, [])

  const refresh = useCallback(() => {
    setEntries(loadNewsletterEntries())
  }, [])

  const create = useCallback((input: CreateInput) => {
    const created = createNewsletterEntry(input)
    setEntries(loadNewsletterEntries())
    return created
  }, [])

  const update = useCallback((id: string, patch: UpdateInput) => {
    const updated = updateNewsletterEntry(id, patch)
    setEntries(loadNewsletterEntries())
    return updated
  }, [])

  const remove = useCallback((id: string) => {
    deleteNewsletterEntry(id)
    setEntries(loadNewsletterEntries())
  }, [])

  const get = useCallback(
    (id: string) => entries.find((entry) => entry.id === id),
    [entries],
  )

  return { entries, refresh, create, update, remove, get }
}
