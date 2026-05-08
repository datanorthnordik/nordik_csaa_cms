import type { EventMedia } from '../api/eventsApi'
import { eventsApi } from '../api/eventsApi'
import { API_ROUTES } from '../constants/api'

function normalizeApiPath(value?: string) {
  if (!value) {
    return null
  }

  if (value.startsWith('/api/')) {
    return value
  }

  if (!/^https?:\/\//i.test(value)) {
    return null
  }

  try {
    const parsedUrl = new URL(value)
    return parsedUrl.pathname.startsWith('/api/')
      ? `${parsedUrl.pathname}${parsedUrl.search}`
      : null
  } catch {
    return null
  }
}

export function resolveEventMediaApiPath(
  media: Pick<EventMedia, 'id' | 'event_id'> &
    Partial<Pick<EventMedia, 'fetch_url' | 'file_url'>>,
) {
  return (
    normalizeApiPath(media.fetch_url) ??
    normalizeApiPath(media.file_url) ??
    API_ROUTES.eventMediaById(media.event_id, media.id)
  )
}

export async function fetchEventMediaObjectUrl(media: EventMedia) {
  const blob = await eventsApi.fetchEventMediaContent(resolveEventMediaApiPath(media))
  return URL.createObjectURL(blob)
}
