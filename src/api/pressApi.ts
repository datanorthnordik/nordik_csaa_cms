import { rawPressReleases } from '../data/pressReleasesData'
import { API_ROUTES } from '../constants/api'
import { apiClient } from './apiClient'
import type { PressEntry, PressStatus, PressVisibility } from '../lib/pressTypes'

// ── Types ─────────────────────────────────────────────────────────────────────

export type PressApiEntry = {
  id: string
  title: string
  release_date: string
  category_id: string | null
  source_url: string
  content_html: string
  status: PressStatus
  visibility: PressVisibility
  publish_at: string | null
  cover_image_url?: string
  created_at: string
  updated_at: string
}

export type PressApiCreateInput = Omit<PressApiEntry, 'id' | 'created_at' | 'updated_at'>

export type PressApiListResponse = {
  items: PressApiEntry[]
  total: number
  page: number
  page_size: number
}

// ── Converters ────────────────────────────────────────────────────────────────

export function pressApiEntryToLocal(entry: PressApiEntry): PressEntry {
  return {
    id: entry.id,
    title: entry.title,
    releaseDate: entry.release_date,
    categoryId: entry.category_id,
    sourceUrl: entry.source_url,
    contentHtml: entry.content_html,
    status: entry.status,
    visibility: entry.visibility,
    publishAt: entry.publish_at,
    coverImageUrl: entry.cover_image_url,
    media: [],
    createdAt: entry.created_at,
    updatedAt: entry.updated_at,
  }
}

export function localEntryToPressApiInput(entry: PressEntry): PressApiCreateInput {
  return {
    title: entry.title,
    release_date: entry.releaseDate,
    category_id: entry.categoryId,
    source_url: entry.sourceUrl,
    content_html: entry.contentHtml,
    status: entry.status,
    visibility: entry.visibility,
    publish_at: entry.publishAt,
    cover_image_url: entry.coverImageUrl,
  }
}

// ── CRUD ──────────────────────────────────────────────────────────────────────

export async function fetchPressEntries(
  params?: Record<string, string | number>,
): Promise<PressApiListResponse> {
  const { data } = await apiClient.get<PressApiListResponse>(API_ROUTES.press, { params })
  return data
}

export async function fetchPressEntry(id: string): Promise<PressApiEntry> {
  const { data } = await apiClient.get<PressApiEntry>(API_ROUTES.pressById(id))
  return data
}

export async function createPressApiEntry(
  input: PressApiCreateInput,
): Promise<PressApiEntry> {
  const { data } = await apiClient.post<PressApiEntry>(API_ROUTES.press, input)
  return data
}

export async function updatePressApiEntry(
  id: string,
  patch: Partial<PressApiCreateInput>,
): Promise<PressApiEntry> {
  const { data } = await apiClient.patch<PressApiEntry>(API_ROUTES.pressById(id), patch)
  return data
}

export async function deletePressApiEntry(id: string): Promise<void> {
  await apiClient.delete(API_ROUTES.pressById(id))
}

// ── Seed utility ──────────────────────────────────────────────────────────────

export type SeedResult = {
  success: number
  failed: number
  errors: Array<{ title: string; error: string }>
}

/**
 * Seeds the backend database with the canonical press releases dataset.
 * Call this once from the browser console (or a migration script) when the
 * backend API is ready:
 *
 *   import { seedPressEntries } from './src/api/pressApi'
 *   const result = await seedPressEntries()
 *   console.log(result)
 */
export async function seedPressEntries(): Promise<SeedResult> {
  const result: SeedResult = { success: 0, failed: 0, errors: [] }

  for (const item of rawPressReleases) {
    const dateIso = `${item.publishDate}T12:00:00.000Z`
    const input: PressApiCreateInput = {
      title: item.title,
      release_date: item.publishDate,
      category_id: null,
      source_url: item.sourceLink,
      content_html: '',
      status: 'published',
      visibility: 'public',
      publish_at: null,
    }

    try {
      await createPressApiEntry(input)
      result.success++
    } catch (error) {
      result.failed++
      result.errors.push({
        title: item.title,
        error: error instanceof Error ? error.message : String(error),
      })
    }

    // Small delay to avoid hammering the API
    await new Promise((resolve) => setTimeout(resolve, 50))
  }

  return result
}
