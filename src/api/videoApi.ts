import { API_ROUTES } from '../constants/api'
import { assertValidGalleryImageUploadFile } from '../lib/resourceUpload'
import { buildMultipartPayload } from './multipartForm'
import { apiClient } from './apiClient'

export type VideoPackageType = 'single' | 'collection'

type ApiVideoPackageSummary = {
  id: number
  title: string
  package_type: VideoPackageType
  video_count: number
  front_image_url?: string
  created_at: string
  updated_at: string
}

type ApiVideoItem = {
  id: number
  video_package_id: number
  title: string
  youtube_url: string
  description: string
  teaser_image_url: string
  storage_uri?: string
  gcp_object_key?: string
  sort_order: number
  created_at: string
  updated_at: string
}

type ApiVideoPackageDetail = {
  id: number
  title: string
  package_type: VideoPackageType
  video_count: number
  single_video?: ApiVideoItem | null
  videos: ApiVideoItem[]
  created_at: string
  updated_at: string
}

type VideoPackageListResponse = {
  items: ApiVideoPackageSummary[]
}

export type VideoPackageSummary = {
  id: number
  title: string
  packageType: VideoPackageType
  videoCount: number
  frontImageUrl?: string
  frontImagePath?: string
  createdAt: string
  updatedAt: string
}

export type VideoItem = {
  id: number
  videoPackageId: number
  title: string
  youtubeUrl: string
  description: string
  teaserImageUrl: string
  teaserImagePath: string
  storageUri?: string
  objectKey?: string
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type VideoPackageDetail = {
  id: number
  title: string
  packageType: VideoPackageType
  videoCount: number
  singleVideo?: VideoItem
  videos: VideoItem[]
  createdAt: string
  updatedAt: string
}

export type VideoUploadInput = {
  title: string
  youtube_url: string
  description?: string
  file_name?: string
  mime_type?: string
  file_url?: string
  storage_uri?: string
  object_key?: string
  gcp_object_key?: string
  remove_teaser_image?: boolean
}

export type CreateVideoPackagePayload = {
  title: string
  package_type: VideoPackageType
  single_video?: VideoUploadInput
  videos?: VideoUploadInput[]
}

export type CreateVideoPackageRequest = CreateVideoPackagePayload & {
  singleVideoFile?: File | null
  videoFiles?: Array<File | null | undefined>
}

export type UpdateVideoPackagePayload = {
  title: string
}

export type UpdateVideoItemPayload = VideoUploadInput

export type UpdateVideoItemRequest = UpdateVideoItemPayload & {
  teaserImageFile?: File | null
}

export type AddVideoItemsPayload = {
  videos: VideoUploadInput[]
}

export type AddVideoItemsRequest = AddVideoItemsPayload & {
  videoFiles?: Array<File | null | undefined>
}

export type VideoPackageMutationResponse = {
  message: string
  video: {
    id: number
    title: string
    package_type: VideoPackageType
  }
}

export type VideoItemsUploadResponse = {
  message: string
  uploadedCount: number
}

export type VideoItemMutationResponse = {
  message: string
  item: VideoItem
}

export type DeleteVideoPackageResponse = {
  message: string
}

export type DeleteVideoItemResponse = {
  message: string
  deletedCount: number
}

function mapVideoPackageSummary(item: ApiVideoPackageSummary): VideoPackageSummary {
  return {
    id: item.id,
    title: item.title,
    packageType: item.package_type,
    videoCount: item.video_count,
    frontImageUrl: item.front_image_url || undefined,
    frontImagePath: item.front_image_url || undefined,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapVideoItem(item: ApiVideoItem): VideoItem {
  return {
    id: item.id,
    videoPackageId: item.video_package_id,
    title: item.title,
    youtubeUrl: item.youtube_url,
    description: item.description || '',
    teaserImageUrl: item.teaser_image_url,
    teaserImagePath: item.teaser_image_url,
    storageUri: item.storage_uri || undefined,
    objectKey: item.gcp_object_key || undefined,
    sortOrder: item.sort_order,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapVideoPackageDetail(detail: ApiVideoPackageDetail): VideoPackageDetail {
  const videos = (detail.videos ?? []).map(mapVideoItem)
  const singleVideo = detail.single_video ? mapVideoItem(detail.single_video) : videos[0]

  return {
    id: detail.id,
    title: detail.title,
    packageType: detail.package_type,
    videoCount: detail.video_count,
    singleVideo,
    videos,
    createdAt: detail.created_at,
    updatedAt: detail.updated_at,
  }
}

function videoItemFileField(index: number) {
  return `videos[${index}].teaser_image_file`
}

function assertValidVideoImage(file?: File | null) {
  if (file) {
    assertValidGalleryImageUploadFile(file)
  }
}

function buildCreateVideoPackageBody(request: CreateVideoPackageRequest) {
  const { singleVideoFile, videoFiles, ...payload } = request

  assertValidVideoImage(singleVideoFile)
  videoFiles?.forEach((file) => assertValidVideoImage(file))

  const hasFiles = Boolean(singleVideoFile) || Boolean(videoFiles?.some(Boolean))
  if (!hasFiles) {
    return payload
  }

  return buildMultipartPayload(payload, [
    ...(singleVideoFile
      ? [
          {
            fieldName: 'single_video.teaser_image_file',
            file: singleVideoFile,
            fileName: singleVideoFile.name,
          },
        ]
      : []),
    ...(videoFiles ?? []).map((file, index) => ({
      fieldName: videoItemFileField(index),
      file,
      fileName: file?.name,
    })),
  ])
}

function buildAddVideoItemsBody(request: AddVideoItemsRequest) {
  const { videoFiles, ...payload } = request

  videoFiles?.forEach((file) => assertValidVideoImage(file))

  if (!videoFiles?.some(Boolean)) {
    return payload
  }

  return buildMultipartPayload(
    payload,
    (videoFiles ?? []).map((file, index) => ({
      fieldName: videoItemFileField(index),
      file,
      fileName: file?.name,
    })),
  )
}

function buildUpdateVideoItemBody(request: UpdateVideoItemRequest) {
  const { teaserImageFile, ...payload } = request

  assertValidVideoImage(teaserImageFile)

  if (!teaserImageFile) {
    return payload
  }

  return buildMultipartPayload(payload, [
    {
      fieldName: 'teaser_image_file',
      file: teaserImageFile,
      fileName: teaserImageFile.name,
    },
  ])
}

export const videoApi = {
  async listVideoPackages() {
    const response = await apiClient.get<VideoPackageListResponse>(API_ROUTES.videos)
    return (response.data.items ?? []).map(mapVideoPackageSummary)
  },

  async getVideoPackage(id: number) {
    const response = await apiClient.get<ApiVideoPackageDetail>(API_ROUTES.videoById(id))
    return mapVideoPackageDetail(response.data)
  },

  async fetchVideoTeaserContent(path: string) {
    const response = await apiClient.get<Blob>(path, {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async createVideoPackage(request: CreateVideoPackageRequest) {
    const response = await apiClient.post<VideoPackageMutationResponse>(
      API_ROUTES.videos,
      buildCreateVideoPackageBody(request),
    )
    return response.data
  },

  async updateVideoPackage(id: number, payload: UpdateVideoPackagePayload) {
    const response = await apiClient.put<VideoPackageMutationResponse>(
      API_ROUTES.videoById(id),
      payload,
    )
    return response.data
  },

  async deleteVideoPackage(id: number) {
    const response = await apiClient.delete<DeleteVideoPackageResponse>(API_ROUTES.videoById(id))
    return response.data
  },

  async addVideoItems(id: number, request: AddVideoItemsRequest) {
    const response = await apiClient.post<VideoItemsUploadResponse>(
      `${API_ROUTES.videoById(id)}/items`,
      buildAddVideoItemsBody(request),
    )
    return response.data
  },

  async updateVideoItem(packageId: number, itemId: number, request: UpdateVideoItemRequest) {
    const response = await apiClient.patch<{ message: string; item: ApiVideoItem }>(
      API_ROUTES.videoItemById(packageId, itemId),
      buildUpdateVideoItemBody(request),
    )

    return {
      message: response.data.message,
      item: mapVideoItem(response.data.item),
    } satisfies VideoItemMutationResponse
  },

  async deleteVideoItem(packageId: number, itemId: number) {
    const response = await apiClient.delete<DeleteVideoItemResponse>(
      API_ROUTES.videoItemById(packageId, itemId),
    )
    return response.data
  },
}
