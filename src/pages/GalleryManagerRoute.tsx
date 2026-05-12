import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useParams } from 'react-router-dom'
import { downloadGalleryContent, fetchGalleryObjectUrl } from '../lib/galleryMedia'
import { fileToBase64 } from '../lib/fileUpload'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import {
  clearCurrentGallery,
  deleteGalleryAssets,
  fetchGalleryById,
  fetchMediaLibrary,
  reorderGalleryAssets,
  saveGallery,
  selectCurrentGallery,
  selectMediaGalleryDetail,
  selectMediaIsSaving,
  updateGalleryAsset,
  uploadGalleryImages,
} from '../store/mediaSlice'
import type { GalleryAsset, GalleryDetail } from '../types/media'
import {
  GalleryManagerPage,
  type GalleryUpdatePatch,
  type PendingUploadInput,
} from './GalleryManagerPage'

const EMPTY_IMAGE_SRC = 'data:,'

export function GalleryManagerRoute() {
  const dispatch = useAppDispatch()
  const { t } = useTranslation()
  const { galleryId } = useParams()
  const detailState = useAppSelector(selectMediaGalleryDetail)
  const rawGallery = useAppSelector(selectCurrentGallery)
  const isSaving = useAppSelector(selectMediaIsSaving)
  const mutationError = useAppSelector(
    (state) =>
      state.media.save.error ??
      state.media.upload.error ??
      state.media.assetUpdate.error ??
      state.media.assetDelete.error ??
      state.media.reorder.error ??
      state.media.deleteGallery.error,
  )
  const [previewUrls, setPreviewUrls] = useState<Record<string, string>>({})

  const numericId = galleryId ? Number.parseInt(galleryId, 10) : Number.NaN

  useEffect(() => {
    if (!Number.isFinite(numericId)) {
      dispatch(clearCurrentGallery())
      return
    }

    void dispatch(fetchGalleryById(numericId))
  }, [dispatch, numericId])

  useEffect(() => {
    if (!rawGallery) {
      setPreviewUrls((current) => {
        Object.values(current).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return {}
      })
      return
    }

    const contentEntries: Array<[string, string]> = []
    if (rawGallery.coverImage?.contentPath) {
      contentEntries.push(['cover', rawGallery.coverImage.contentPath])
    }
    for (const asset of rawGallery.assets ?? []) {
      const contentPath = asset.contentPath || asset.fileUrl
      if (contentPath) {
        contentEntries.push([String(asset.id), contentPath])
      }
    }

    let cancelled = false

    async function loadPreviews() {
      const loadedEntries = await Promise.all(
        contentEntries.map(async ([key, contentPath]) => {
          try {
            const objectUrl = await fetchGalleryObjectUrl(contentPath)
            return [key, objectUrl] as const
          } catch {
            return null
          }
        }),
      )

      const nextUrls = loadedEntries.reduce<Record<string, string>>((result, entry) => {
        if (entry) {
          result[entry[0]] = entry[1]
        }
        return result
      }, {})

      if (cancelled) {
        Object.values(nextUrls).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return
      }

      setPreviewUrls((current) => {
        Object.values(current).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return nextUrls
      })
    }

    void loadPreviews()

    return () => {
      cancelled = true
    }
  }, [rawGallery])

  const gallery = useMemo(() => {
    if (!rawGallery) {
      return undefined
    }

    return {
      ...rawGallery,
      coverImage: rawGallery.coverImage
        ? {
            ...rawGallery.coverImage,
            fileUrl: previewUrls.cover ?? EMPTY_IMAGE_SRC,
          }
        : undefined,
      assets: (rawGallery.assets ?? []).map((asset) => ({
        ...asset,
        fileUrl: previewUrls[String(asset.id)] ?? EMPTY_IMAGE_SRC,
      })),
    } satisfies GalleryDetail
  }, [previewUrls, rawGallery])

  async function refreshGallery(options: { refreshList?: boolean } = {}) {
    if (!Number.isFinite(numericId)) {
      return
    }

    await dispatch(fetchGalleryById(numericId)).unwrap()
    if (options.refreshList) {
      await dispatch(fetchMediaLibrary()).unwrap()
    }
  }

  async function persistGalleryPatch(
    patch: GalleryUpdatePatch,
    options: {
      published?: boolean
      coverFile?: File | null
      removeCover?: boolean
      toastMessage?: string
      refreshList?: boolean
    } = {},
  ) {
    if (!rawGallery || !Number.isFinite(numericId)) {
      return false
    }

    try {
      const payload = await buildSaveGalleryPayload(rawGallery, patch, options)
      const result = await dispatch(
        saveGallery({
          id: numericId,
          payload,
        }),
      ).unwrap()

      await refreshGallery({ refreshList: options.refreshList ?? true })

      if (options.toastMessage) {
        toast.success(options.toastMessage)
      } else if (options.coverFile !== undefined || options.removeCover) {
        toast.success(result.message)
      }

      return true
    } catch {
      return false
    }
  }

  async function handleUploadAssets(uploads: PendingUploadInput[]) {
    if (!rawGallery || !Number.isFinite(numericId)) {
      return
    }

    try {
      const payload = {
        images: await Promise.all(
          uploads.map(async (upload) => ({
            title: upload.title.trim(),
            alt_text: upload.details.trim(),
            file_name: upload.file.name,
            mime_type: upload.file.type || 'application/octet-stream',
            data_base64: await fileToBase64(upload.file),
          })),
        ),
      }

      const result = await dispatch(
        uploadGalleryImages({
          id: numericId,
          payload,
        }),
      ).unwrap()

      await refreshGallery({ refreshList: true })
      toast.success(result.message)
    } catch {
      return
    }
  }

  async function handleUpdateAsset(asset: GalleryAsset, patch: { title: string; details: string }) {
    if (!Number.isFinite(numericId)) {
      return
    }

    try {
      const result = await dispatch(
        updateGalleryAsset({
          galleryId: numericId,
          imageId: asset.id,
          payload: {
            title: patch.title.trim(),
            alt_text: patch.details.trim(),
          },
        }),
      ).unwrap()

      await refreshGallery()
      toast.success(result.message)
    } catch {
      return
    }
  }

  async function handleDeleteAsset(asset: GalleryAsset) {
    if (!asset.storageUri || !Number.isFinite(numericId)) {
      return
    }

    try {
      const result = await dispatch(
        deleteGalleryAssets({
          id: numericId,
          payload: {
            storage_urls: [asset.storageUri],
          },
        }),
      ).unwrap()

      await refreshGallery({ refreshList: true })
      toast.success(result.message)
    } catch {
      return
    }
  }

  async function persistImageOrder(nextImageIds: number[]) {
    if (!Number.isFinite(numericId)) {
      return
    }

    try {
      await dispatch(
        reorderGalleryAssets({
          id: numericId,
          payload: {
            image_ids: nextImageIds,
          },
        }),
      ).unwrap()

      await refreshGallery()
    } catch {
      return
    }
  }

  async function handleMoveAsset(asset: GalleryAsset, delta: number) {
    const assets = rawGallery?.assets ?? []
    const fromIndex = assets.findIndex((item) => item.id === asset.id)
    const toIndex = fromIndex + delta

    if (fromIndex === -1 || toIndex < 0 || toIndex >= assets.length) {
      return
    }

    const next = [...assets]
    const [moved] = next.splice(fromIndex, 1)
    next.splice(toIndex, 0, moved)
    await persistImageOrder(next.map((item) => item.id))
  }

  async function handleReorderAsset(fromAsset: GalleryAsset, toAsset: GalleryAsset) {
    const assets = rawGallery?.assets ?? []
    const fromIndex = assets.findIndex((item) => item.id === fromAsset.id)
    const toIndex = assets.findIndex((item) => item.id === toAsset.id)

    if (fromIndex === -1 || toIndex === -1 || fromIndex === toIndex) {
      return
    }

    const next = [...assets]
    const [moved] = next.splice(fromIndex, 1)
    next.splice(toIndex, 0, moved)
    await persistImageOrder(next.map((item) => item.id))
  }

  async function handleDownloadAsset(asset: GalleryAsset) {
    const contentPath = asset.contentPath || rawGallery?.assets?.find((item) => item.id === asset.id)?.contentPath
    if (!contentPath) {
      return
    }

    try {
      await downloadGalleryContent(contentPath, asset.fileName)
    } catch {
      toast.error(t('galleryManager.feedback.downloadFailed'))
    }
  }

  return (
    <GalleryManagerPage
      gallery={gallery}
      loading={detailState.status === 'loading'}
      saving={isSaving}
      error={detailState.error ?? mutationError ?? undefined}
      onUpdateGallery={(patch) => {
        void persistGalleryPatch(patch)
      }}
      onUploadAssets={(uploads) => {
        void handleUploadAssets(uploads)
      }}
      onUpdateAsset={(asset, patch) => {
        void handleUpdateAsset(asset, patch)
      }}
      onDeleteAsset={(asset) => {
        void handleDeleteAsset(asset)
      }}
      onMoveAsset={(asset, delta) => {
        void handleMoveAsset(asset, delta)
      }}
      onReorderAsset={(fromAsset, toAsset) => {
        void handleReorderAsset(fromAsset, toAsset)
      }}
      onSetCover={(file) => {
        void persistGalleryPatch(
          {},
          {
            coverFile: file,
            removeCover: file === null,
            refreshList: true,
          },
        )
      }}
      onDownloadAsset={(asset) => {
        void handleDownloadAsset(asset)
      }}
      onSaveDraft={() => {
        void persistGalleryPatch(
          {},
          {
            published: false,
            refreshList: true,
            toastMessage: t('galleryManager.feedback.draftSaved'),
          },
        )
      }}
      onPublish={() => {
        void persistGalleryPatch(
          {},
          {
            published: true,
            refreshList: true,
            toastMessage: t('galleryManager.feedback.published'),
          },
        )
      }}
    />
  )
}

async function buildSaveGalleryPayload(
  gallery: GalleryDetail,
  patch: GalleryUpdatePatch,
  options: {
    published?: boolean
    coverFile?: File | null
    removeCover?: boolean
  } = {},
) {
  const name = (patch.name ?? gallery.name).trim()
  const descriptionValue =
    patch.description !== undefined
      ? patch.description
      : (gallery.description ?? '')

  return {
    name,
    description: descriptionValue.trim() || undefined,
    published:
      options.published !== undefined
        ? options.published
        : gallery.visibility === 'published',
    cover_image: options.coverFile
      ? {
          file_name: options.coverFile.name,
          mime_type: options.coverFile.type || 'application/octet-stream',
          data_base64: await fileToBase64(options.coverFile),
          alt_text: gallery.coverImage?.altText || name,
        }
      : undefined,
    remove_cover_image: options.removeCover ?? false,
  }
}
