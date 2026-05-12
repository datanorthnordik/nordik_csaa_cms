import type { GalleryAsset } from '../types/media'

const fileExtensionPattern = /\.[^.]+$/

export function suggestGalleryAssetTitle(fileName: string) {
  const suggested = fileName
    .replace(fileExtensionPattern, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

  return suggested || fileName
}

export function getGalleryAssetTitle(
  asset: Pick<GalleryAsset, 'fileName' | 'title'>,
) {
  const title = asset.title?.trim()
  return title || suggestGalleryAssetTitle(asset.fileName)
}

export function getGalleryAssetDetails(
  asset: Pick<GalleryAsset, 'details' | 'altText'>,
) {
  return asset.details?.trim() || asset.altText?.trim() || ''
}
