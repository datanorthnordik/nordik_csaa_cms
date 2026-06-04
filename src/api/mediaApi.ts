import { API_ROUTES } from '../constants/api'
import { assertValidGalleryImageUploadFile } from '../lib/resourceUpload'
import type { GalleryAsset, GalleryDetail, GallerySummary } from '../types/media'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

type ApiGallerySummary = {
  id: number
  name: string
  published: boolean
  asset_count: number
  front_image_url?: string
  created_at: string
  updated_at: string
}

type ApiGalleryAsset = {
  id: number
  gallery_id: number
  title: string
  alt_text: string
  link_url?: string
  file_name: string
  gcp_object_key?: string
  file_url: string
  storage_uri?: string
  mime_type?: string
  file_size?: number
  sort_order?: number
  uploaded_by?: number | null
  created_at: string
  updated_at: string
}

type ApiGalleryDetail = {
  id: number
  name: string
  description: string
  published: boolean
  asset_limit: number
  cover_image?: ApiGalleryAsset | null
  images: ApiGalleryAsset[]
  created_by?: number | null
  updated_by?: number | null
  created_at: string
  updated_at: string
}

type GalleryListResponse = {
  items: ApiGallerySummary[]
}

type GalleryDetailEnvelope = ApiGalleryDetail

export type GalleryUploadInput = {
  title?: string
  alt_text?: string
  link_url?: string
  file_name?: string
  mime_type?: string
  file_url?: string
  storage_uri?: string
  object_key?: string
  gcp_object_key?: string
}

export type SaveGalleryPayload = {
  name: string
  description?: string
  published: boolean
  cover_image?: GalleryUploadInput
  remove_cover_image?: boolean
}

export type SaveGalleryRequest = SaveGalleryPayload & {
  coverImageFile?: File
}

export type CreateGalleryResponse = {
  message: string
  gallery: {
    id: number
    name: string
    published: boolean
  }
}

export type UpdateGalleryResponse = CreateGalleryResponse

export type UploadGalleryImagesPayload = {
  images: GalleryUploadInput[]
}

export type UploadGalleryImagesRequest = UploadGalleryImagesPayload & {
  imageFiles?: Array<File | null | undefined>
}

export type UploadGalleryImagesResponse = {
  message: string
  uploadedCount: number
}

export type UpdateGalleryImagePayload = {
  title: string
  alt_text: string
  link_url?: string
}

export type UpdateGalleryImageResponse = {
  message: string
  image: GalleryAsset
}

export type ReorderGalleryImagesPayload = {
  image_ids: number[]
}

export type ReorderGalleryImagesResponse = {
  message: string
  updatedCount: number
}

export type DeleteGalleryImagesPayload = {
  storage_urls: string[]
}

export type DeleteGalleryImagesResponse = {
  message: string
  deletedCount: number
}

export type DeleteGalleryResponse = {
  message: string
}

function toVisibility(published: boolean) {
  return published ? 'published' : 'draft'
}

function mapGalleryAsset(asset: ApiGalleryAsset): GalleryAsset {
  return {
    id: asset.id,
    fileName: asset.file_name || asset.title || `gallery-image-${asset.id}`,
    fileUrl: asset.file_url,
    contentPath: asset.file_url,
    storageUri: asset.storage_uri,
    objectKey: asset.gcp_object_key,
    mimeType: asset.mime_type,
    fileSize: asset.file_size,
    sortOrder: asset.sort_order,
    title: asset.title || undefined,
    details: asset.alt_text || undefined,
    altText: asset.alt_text || undefined,
    linkUrl: asset.link_url || undefined,
    uploadedAt: asset.created_at,
  }
}

function mapGallerySummary(item: ApiGallerySummary): GallerySummary {
  return {
    id: item.id,
    name: item.name,
    assetCount: item.asset_count,
    frontImageUrl: item.front_image_url,
    frontImagePath: item.front_image_url,
    visibility: toVisibility(item.published),
    updatedAt: item.updated_at,
  }
}

function mapGalleryDetail(detail: ApiGalleryDetail): GalleryDetail {
  return {
    id: detail.id,
    name: detail.name,
    description: detail.description || undefined,
    assetLimit: detail.asset_limit,
    visibility: toVisibility(detail.published),
    updatedAt: detail.updated_at,
    coverImage: detail.cover_image ? mapGalleryAsset(detail.cover_image) : undefined,
    assets: (detail.images ?? []).map(mapGalleryAsset),
  }
}

export const mediaApi = {
  async listGalleries() {
    const response = await apiClient.get<GalleryListResponse>(API_ROUTES.galleries)
    return response.data.items.map(mapGallerySummary)
  },

  async getGallery(id: number) {
    const response = await apiClient.get<GalleryDetailEnvelope>(API_ROUTES.galleryById(id))
    return mapGalleryDetail(response.data)
  },

  async fetchGalleryContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createGallery(request: SaveGalleryRequest) {
    const { coverImageFile, ...payload } = request
    if (coverImageFile) {
      assertValidGalleryImageUploadFile(coverImageFile)
    }

    const body = coverImageFile
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'cover_image_file',
            file: coverImageFile,
            fileName: coverImageFile.name,
          },
        ])
      : payload
    const response = await apiClient.post<CreateGalleryResponse>(
      API_ROUTES.galleries,
      body,
    )
    return response.data
  },

  async updateGallery(id: number, request: SaveGalleryRequest) {
    const { coverImageFile, ...payload } = request
    if (coverImageFile) {
      assertValidGalleryImageUploadFile(coverImageFile)
    }

    const body = coverImageFile
      ? buildMultipartPayload(payload, [
          {
            fieldName: 'cover_image_file',
            file: coverImageFile,
            fileName: coverImageFile.name,
          },
        ])
      : payload
    const response = await apiClient.put<UpdateGalleryResponse>(
      API_ROUTES.galleryById(id),
      body,
    )
    return response.data
  },

  async deleteGallery(id: number) {
    const response = await apiClient.delete<DeleteGalleryResponse>(
      API_ROUTES.galleryById(id),
    )
    return response.data
  },

  async uploadGalleryImages(id: number, request: UploadGalleryImagesRequest) {
    const { imageFiles, ...payload } = request
    imageFiles?.forEach((file) => {
      if (file) {
        assertValidGalleryImageUploadFile(file)
      }
    })

    const hasFiles = Boolean(imageFiles?.some(Boolean))
    const body = hasFiles
      ? buildMultipartPayload(
          payload,
          (imageFiles ?? []).map((file, index) => ({
            fieldName: `images[${index}].file`,
            file,
            fileName: file?.name,
          })),
        )
      : payload
    const response = await apiClient.post<UploadGalleryImagesResponse>(
      API_ROUTES.galleryImagesById(id),
      body,
    )
    return response.data
  },

  async updateGalleryImage(
    galleryId: number,
    imageId: number,
    payload: UpdateGalleryImagePayload,
  ) {
    const response = await apiClient.patch<{ message: string; image: ApiGalleryAsset }>(
      API_ROUTES.galleryImageById(galleryId, imageId),
      payload,
    )

    return {
      message: response.data.message,
      image: mapGalleryAsset(response.data.image),
    } satisfies UpdateGalleryImageResponse
  },

  async reorderGalleryImages(id: number, payload: ReorderGalleryImagesPayload) {
    const response = await apiClient.put<ReorderGalleryImagesResponse>(
      API_ROUTES.galleryImageOrderById(id),
      payload,
    )
    return response.data
  },

  async deleteGalleryImages(id: number, payload: DeleteGalleryImagesPayload) {
    const response = await apiClient.delete<DeleteGalleryImagesResponse>(
      API_ROUTES.galleryImagesById(id),
      {
        data: payload,
      },
    )
    return response.data
  },
}
