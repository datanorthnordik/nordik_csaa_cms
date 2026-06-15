import { describe, expect, it } from 'vitest'
import { extractYouTubeVideoId, getYouTubeEmbedUrl } from './videoPreview'

describe('videoPreview', () => {
  it('extracts ids from common YouTube URL formats', () => {
    expect(extractYouTubeVideoId('https://www.youtube.com/watch?v=welcome123')).toBe(
      'welcome123',
    )
    expect(extractYouTubeVideoId('https://youtu.be/story-one?t=42')).toBe('story-one')
    expect(extractYouTubeVideoId('youtube.com/embed/community-clip')).toBe(
      'community-clip',
    )
    expect(extractYouTubeVideoId('https://m.youtube.com/shorts/youth-story')).toBe(
      'youth-story',
    )
  })

  it('builds embeddable YouTube preview URLs', () => {
    expect(getYouTubeEmbedUrl('https://www.youtube.com/watch?v=welcome123')).toBe(
      'https://www.youtube.com/embed/welcome123?rel=0',
    )
    expect(getYouTubeEmbedUrl('https://youtu.be/story-one')).toBe(
      'https://www.youtube.com/embed/story-one?rel=0',
    )
  })

  it('returns null for empty or non-YouTube URLs', () => {
    expect(extractYouTubeVideoId('')).toBeNull()
    expect(extractYouTubeVideoId('https://example.com/watch?v=welcome123')).toBeNull()
    expect(getYouTubeEmbedUrl('not a valid link')).toBeNull()
  })
})
