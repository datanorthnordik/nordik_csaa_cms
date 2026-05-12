import { mediaApi } from '../api/mediaApi'

export async function fetchGalleryObjectUrl(contentPath: string) {
  const blob = await mediaApi.fetchGalleryContent(contentPath)
  return URL.createObjectURL(blob)
}

export async function downloadGalleryContent(
  contentPath: string,
  fileName: string,
) {
  const blob = await mediaApi.fetchGalleryContent(contentPath)
  const objectUrl = URL.createObjectURL(blob)

  try {
    const link = document.createElement('a')
    link.href = objectUrl
    link.download = fileName
    link.click()
  } finally {
    setTimeout(() => URL.revokeObjectURL(objectUrl), 0)
  }
}
