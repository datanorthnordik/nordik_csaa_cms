import type { NewsletterEntry } from './newsletterTypes'

const STORAGE_KEY = 'nordik_csaa_newsletter_entries:v2'

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
  return `newsletter-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function isValidEntry(value: unknown): value is NewsletterEntry {
  if (!value || typeof value !== 'object') {
    return false
  }
  const entry = value as Partial<NewsletterEntry>
  return (
    typeof entry.id === 'string' &&
    typeof entry.title === 'string' &&
    typeof entry.sendDate === 'string' &&
    typeof entry.status === 'string'
  )
}

function readRaw(): NewsletterEntry[] | null {
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

function writeRaw(entries: NewsletterEntry[]) {
  if (!isBrowser()) {
    return
  }
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(entries))
}

const seedEntries: NewsletterEntry[] = [
  {
    id: 'newsletter-seed-001',
    title: 'Spring 2024 Community Update',
    category: 'csaa',
    sendDate: '2024-03-15',
    contentHtml: '<p>Welcome to the Spring 2024 edition of the CSAA Newsletter.</p>',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2024-03-10T09:00:00.000Z',
    updatedAt: '2024-03-15T12:00:00.000Z',
  },
  {
    id: 'newsletter-seed-002',
    title: 'Summer 2024 — Annual Gathering Preview',
    category: 'csaa',
    sendDate: '2024-06-20',
    contentHtml: '<p>The Annual Gathering is approaching — read on for details.</p>',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2024-06-10T09:00:00.000Z',
    updatedAt: '2024-06-18T14:30:00.000Z',
  },
  {
    id: 'newsletter-seed-003',
    title: 'Fall 2024 Member Highlights',
    category: 'cst',
    sendDate: '2024-09-12',
    contentHtml: '<p>This issue celebrates members who have made a difference.</p>',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2024-09-05T10:00:00.000Z',
    updatedAt: '2024-09-10T11:00:00.000Z',
  },
  {
    id: 'newsletter-seed-004',
    title: 'Winter 2024 Year in Review',
    category: 'cst',
    sendDate: '2024-12-05',
    contentHtml: '<p>Looking back on a remarkable year for our community.</p>',
    status: 'draft',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2024-11-20T08:00:00.000Z',
    updatedAt: '2024-11-25T16:00:00.000Z',
  },
]

export function loadNewsletterEntries(): NewsletterEntry[] {
  const existing = readRaw()
  if (existing) {
    return existing
  }
  writeRaw(seedEntries)
  return seedEntries
}

export function saveNewsletterEntries(entries: NewsletterEntry[]) {
  writeRaw(entries)
}

export function getNewsletterEntry(id: string): NewsletterEntry | undefined {
  return loadNewsletterEntries().find((entry) => entry.id === id)
}

export function createNewsletterEntry(
  input: Omit<NewsletterEntry, 'id' | 'createdAt' | 'updatedAt'>,
): NewsletterEntry {
  const entries = loadNewsletterEntries()
  const timestamp = nowIso()
  const created: NewsletterEntry = {
    ...input,
    id: makeId(),
    createdAt: timestamp,
    updatedAt: timestamp,
  }
  writeRaw([created, ...entries])
  return created
}

export function updateNewsletterEntry(
  id: string,
  patch: Partial<Omit<NewsletterEntry, 'id' | 'createdAt'>>,
): NewsletterEntry | undefined {
  const entries = loadNewsletterEntries()
  const index = entries.findIndex((entry) => entry.id === id)
  if (index === -1) {
    return undefined
  }
  const merged: NewsletterEntry = {
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

export function deleteNewsletterEntry(id: string) {
  const entries = loadNewsletterEntries()
  writeRaw(entries.filter((entry) => entry.id !== id))
}

export const NEWSLETTER_ENTRIES_STORAGE_KEY = STORAGE_KEY
