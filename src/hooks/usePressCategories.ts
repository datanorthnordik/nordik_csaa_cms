import { useMemo } from 'react'
import { loadPressCategories } from '../lib/pressCategoriesStorage'

export function usePressCategories() {
  const categories = useMemo(() => loadPressCategories(), [])
  return { categories }
}
