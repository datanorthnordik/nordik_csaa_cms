import type { PressCategory } from './pressTypes'

const STORAGE_KEY = 'nordik_csaa_press_categories:v1'

function isBrowser() {
  return typeof window !== 'undefined'
}

function nowIso() {
  return new Date().toISOString()
}

function makeId() {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }
  return `cat-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function isValidCategory(value: unknown): value is PressCategory {
  if (!value || typeof value !== 'object') {
    return false
  }
  const cat = value as Partial<PressCategory>
  return typeof cat.id === 'string' && typeof cat.name === 'string'
}

function readRaw(): PressCategory[] | null {
  if (!isBrowser()) {
    return null
  }
  const raw = window.localStorage.getItem(STORAGE_KEY)
  if (!raw) {
    return null
  }
  try {
    const parsed = JSON.parse(raw)
    if (!Array.isArray(parsed)) {
      return null
    }
    return parsed.filter(isValidCategory)
  } catch {
    return null
  }
}

function writeRaw(categories: PressCategory[]) {
  if (!isBrowser()) {
    return
  }
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(categories))
}

const seedCategories: PressCategory[] = [
  { id: 'cat-statement', name: 'Statement', createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'cat-news', name: 'News', createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'cat-announcement', name: 'Announcement', createdAt: '2024-01-01T00:00:00.000Z' },
  {
    id: 'cat-community-outreach',
    name: 'Community Outreach',
    createdAt: '2024-01-01T00:00:00.000Z',
  },
]

export function loadPressCategories(): PressCategory[] {
  const existing = readRaw()
  if (existing) {
    return existing
  }
  writeRaw(seedCategories)
  return seedCategories
}

export function createPressCategory(name: string): PressCategory {
  const categories = loadPressCategories()
  const trimmed = name.trim()
  const existing = categories.find(
    (cat) => cat.name.toLowerCase() === trimmed.toLowerCase(),
  )
  if (existing) {
    return existing
  }
  const created: PressCategory = {
    id: makeId(),
    name: trimmed,
    createdAt: nowIso(),
  }
  writeRaw([...categories, created])
  return created
}

export function renamePressCategory(id: string, name: string): PressCategory | undefined {
  const categories = loadPressCategories()
  const index = categories.findIndex((cat) => cat.id === id)
  if (index === -1) {
    return undefined
  }
  const next = [...categories]
  next[index] = { ...next[index], name: name.trim() }
  writeRaw(next)
  return next[index]
}

export function deletePressCategory(id: string) {
  const categories = loadPressCategories()
  writeRaw(categories.filter((cat) => cat.id !== id))
}

export const PRESS_CATEGORIES_STORAGE_KEY = STORAGE_KEY
