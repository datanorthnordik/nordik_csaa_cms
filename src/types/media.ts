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
  contentPath?: string
  storageUri?: string
  objectKey?: string
  thumbnailUrl?: string
  mimeType?: string
  fileSize?: number
  sortOrder?: number
  dimensions?: GalleryAssetDimensions
  title?: string
  details?: string
  altText?: string
  linkUrl?: string
  uploadedBy?: string
  uploadedAt?: string
  usageTracking?: GalleryAssetUsage[]
}

export type GalleryAssetContentPatch = {
  title: string
  details: string
  linkUrl?: string
}

export type GallerySummary = {
  id: number
  name: string
  assetCount?: number
  frontImageUrl?: string
  frontImagePath?: string
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
