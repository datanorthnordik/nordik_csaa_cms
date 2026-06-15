function parseYouTubeUrl(value: string) {
  const trimmedValue = value.trim()
  if (!trimmedValue) {
    return null
  }

  const normalizedValue = /^[a-zA-Z][a-zA-Z\d+\-.]*:/.test(trimmedValue)
    ? trimmedValue
    : `https://${trimmedValue}`

  try {
    return new URL(normalizedValue)
  } catch {
    return null
  }
}

export function extractYouTubeVideoId(value: string) {
  const parsedUrl = parseYouTubeUrl(value)
  if (!parsedUrl) {
    return null
  }

  const hostname = parsedUrl.hostname.toLowerCase().replace(/^www\./, '')
  if (hostname === 'youtu.be') {
    return parsedUrl.pathname.split('/').filter(Boolean)[0] ?? null
  }

  if (!hostname.endsWith('youtube.com')) {
    return null
  }

  if (parsedUrl.pathname === '/watch') {
    return parsedUrl.searchParams.get('v')?.trim() || null
  }

  const pathSegments = parsedUrl.pathname.split('/').filter(Boolean)
  const previewSegmentIndex = pathSegments.findIndex((segment) =>
    ['embed', 'live', 'shorts', 'v'].includes(segment),
  )

  if (previewSegmentIndex >= 0) {
    return pathSegments[previewSegmentIndex + 1] ?? null
  }

  return null
}

export function getYouTubeEmbedUrl(value: string) {
  const videoId = extractYouTubeVideoId(value)
  if (!videoId) {
    return null
  }

  return `https://www.youtube.com/embed/${encodeURIComponent(videoId)}?rel=0`
}
