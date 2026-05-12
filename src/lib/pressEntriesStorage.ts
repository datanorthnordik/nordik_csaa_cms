import { rawPressReleases } from '../data/pressReleasesData'
import type { PressEntry } from './pressTypes'

const STORAGE_KEY = 'nordik_csaa_press_entries:v2'

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
  return `press-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function isValidEntry(value: unknown): value is PressEntry {
  if (!value || typeof value !== 'object') {
    return false
  }
  const entry = value as Partial<PressEntry>
  return (
    typeof entry.id === 'string' &&
    typeof entry.title === 'string' &&
    typeof entry.releaseDate === 'string' &&
    typeof entry.status === 'string'
  )
}

function readRaw(): PressEntry[] | null {
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
    return parsed.filter(isValidEntry)
  } catch {
    return null
  }
}

function writeRaw(entries: PressEntry[]) {
  if (!isBrowser()) {
    return
  }
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(entries))
}

const seedEntries: PressEntry[] = rawPressReleases.map((item, index) => {
  const paddedIndex = String(index + 1).padStart(3, '0')
  const dateIso = `${item.publishDate}T12:00:00.000Z`
  return {
    id: `press-seed-${paddedIndex}`,
    title: item.title,
    releaseDate: item.publishDate,
    categoryId: null,
    sourceUrl: item.sourceLink,
    contentHtml: '',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: dateIso,
    updatedAt: dateIso,
  }
})

export function loadPressEntries(): PressEntry[] {
  const existing = readRaw()
  if (existing) {
    return existing
  }
  writeRaw(seedEntries)
  return seedEntries
}

export function savePressEntries(entries: PressEntry[]) {
  writeRaw(entries)
}

export function getPressEntry(id: string): PressEntry | undefined {
  return loadPressEntries().find((entry) => entry.id === id)
}

export function createPressEntry(
  input: Omit<PressEntry, 'id' | 'createdAt' | 'updatedAt'>,
): PressEntry {
  const entries = loadPressEntries()
  const timestamp = nowIso()
  const created: PressEntry = {
    ...input,
    id: makeId(),
    createdAt: timestamp,
    updatedAt: timestamp,
  }
  writeRaw([created, ...entries])
  return created
}

export function updatePressEntry(
  id: string,
  patch: Partial<Omit<PressEntry, 'id' | 'createdAt'>>,
): PressEntry | undefined {
  const entries = loadPressEntries()
  const index = entries.findIndex((entry) => entry.id === id)
  if (index === -1) {
    return undefined
  }
  const merged: PressEntry = {
    ...entries[index],
    ...patch,
    id: entries[index].id,
    createdAt: entries[index].createdAt,
    updatedAt: nowIso(),
  }
  const next = [...entries]
  next[index] = merged
  writeRaw(next)
  return merged
}

export function deletePressEntry(id: string) {
  const entries = loadPressEntries()
  writeRaw(entries.filter((entry) => entry.id !== id))
}

export const PRESS_ENTRIES_STORAGE_KEY = STORAGE_KEY
