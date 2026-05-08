import { describe, expect, it } from 'vitest'
import type { EventMedia } from '../api/eventsApi'
import { resolveEventMediaApiPath } from './eventMedia'

function createMedia(overrides: Partial<EventMedia> = {}): EventMedia {
  return {
    id: 23,
    event_id: 17,
    media_role: 'attachment',
    display_name: 'diagram.png',
    gcp_object_key: 'events/17/diagram.png',
    file_url:
      'https://storage.googleapis.com/nordik-csa-documents/events/17/diagram.png',
    mime_type: 'image/png',
    file_size: 1024,
    sort_order: 0,
    created_at: '2026-05-07T10:00:00Z',
    updated_at: '2026-05-07T10:00:00Z',
    ...overrides,
  }
}

describe('resolveEventMediaApiPath', () => {
  it('falls back to the authenticated event media API for public GCS urls', () => {
    expect(resolveEventMediaApiPath(createMedia())).toBe(
      '/api/events/17/media/23/content',
    )
  })

  it('uses a provided authenticated fetch url when available', () => {
    expect(
      resolveEventMediaApiPath(
        createMedia({
          fetch_url: '/api/events/17/media/23/content',
        }),
      ),
    ).toBe('/api/events/17/media/23/content')
  })
})
