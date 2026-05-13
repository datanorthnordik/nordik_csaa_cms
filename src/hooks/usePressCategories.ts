import { useCallback, useEffect, useState } from 'react'
import {
  createPressCategory,
  deletePressCategory,
  loadPressCategories,
  PRESS_CATEGORIES_STORAGE_KEY,
  renamePressCategory,
} from '../lib/pressCategoriesStorage'
import type { PressCategory } from '../lib/pressTypes'

export function usePressCategories() {
  const [categories, setCategories] = useState<PressCategory[]>(() =>
    loadPressCategories(),
  )

  useEffect(() => {
    function handleStorage(event: StorageEvent) {
      if (event.key !== PRESS_CATEGORIES_STORAGE_KEY) {
        return
      }
      setCategories(loadPressCategories())
    }
    window.addEventListener('storage', handleStorage)
    return () => window.removeEventListener('storage', handleStorage)
  }, [])

  const create = useCallback((name: string) => {
    const created = createPressCategory(name)
    setCategories(loadPressCategories())
    return created
  }, [])

  const rename = useCallback((id: string, name: string) => {
    const renamed = renamePressCategory(id, name)
    setCategories(loadPressCategories())
    return renamed
  }, [])

  const remove = useCallback((id: string) => {
    deletePressCategory(id)
    setCategories(loadPressCategories())
  }, [])

  return { categories, create, rename, remove }
}
