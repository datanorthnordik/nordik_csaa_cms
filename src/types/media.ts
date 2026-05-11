export type MediaVisibility = 'draft' | 'published' | 'public' | 'internal'

export type GalleryAssetUsage = {
  label: string
  path: string
}

export type GalleryAssetDimensions = {
  width: number
  height: number
}

export type GalleryAsset = {
  id: number
  fileName: string
  fileUrl: string
  thumbnailUrl?: string
  mimeType?: string
  fileSize?: number
  dimensions?: GalleryAssetDimensions
  altText?: string
  uploadedBy?: string
  uploadedAt?: string
  usageTracking?: GalleryAssetUsage[]
}

export type GallerySummary = {
  id: number
  name: string
  assetCount?: number
  frontImageUrl?: string
  visibility?: MediaVisibility
  updatedAt?: string
}

export type GalleryDetail = {
  id: number
  name: string
  description?: string
  assetLimit?: number
  visibility?: MediaVisibility
  updatedAt?: string
  coverImage?: GalleryAsset
  assets?: GalleryAsset[]
}
