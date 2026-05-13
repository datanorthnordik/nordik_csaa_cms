import { useCallback, useEffect, useState } from 'react'
import {
  createPressEntry,
  deletePressEntry,
  loadPressEntries,
  PRESS_ENTRIES_STORAGE_KEY,
  updatePressEntry,
} from '../lib/pressEntriesStorage'
import type { PressEntry } from '../lib/pressTypes'

type CreateInput = Omit<PressEntry, 'id' | 'createdAt' | 'updatedAt'>
type UpdateInput = Partial<Omit<PressEntry, 'id' | 'createdAt'>>

export function usePressEntries() {
  const [entries, setEntries] = useState<PressEntry[]>(() => loadPressEntries())

  useEffect(() => {
    function handleStorage(event: StorageEvent) {
      if (event.key !== PRESS_ENTRIES_STORAGE_KEY) {
        return
      }
      setEntries(loadPressEntries())
    }
    window.addEventListener('storage', handleStorage)
    return () => window.removeEventListener('storage', handleStorage)
  }, [])

  const refresh = useCallback(() => {
    setEntries(loadPressEntries())
  }, [])

  const create = useCallback((input: CreateInput) => {
    const created = createPressEntry(input)
    setEntries(loadPressEntries())
    return created
  }, [])

  const update = useCallback((id: string, patch: UpdateInput) => {
    const updated = updatePressEntry(id, patch)
    setEntries(loadPressEntries())
    return updated
  }, [])

  const remove = useCallback((id: string) => {
    deletePressEntry(id)
    setEntries(loadPressEntries())
  }, [])

  const get = useCallback(
    (id: string) => entries.find((entry) => entry.id === id),
    [entries],
  )

  return { entries, refresh, create, update, remove, get }
}
