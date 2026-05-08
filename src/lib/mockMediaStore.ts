import { useCallback, useEffect, useState } from 'react'
import type {
  GalleryAsset,
  GalleryDetail,
  GallerySummary,
  MediaVisibility,
} from '../types/media'

// Temporary in-memory store used to exercise the Media Library / Gallery Manager
// UI before the real API exists. Module-level so multiple pages see the same data
// during a single session. Reload clears state. Replace with a Redux slice +
// mediaApi once the backend is available.

const galleries = new Map<number, GalleryDetail>()
const listeners = new Set<() => void>()
let nextId = 1

function notify() {
  listeners.forEach((fn) => fn())
}

function pickFrontImageUrl(detail: GalleryDetail): string | undefined {
  return detail.assets?.[0]?.fileUrl
}

function summarize(detail: GalleryDetail): GallerySummary {
  return {
    id: detail.id,
    name: detail.name,
    assetCount: detail.assets?.length ?? 0,
    frontImageUrl: pickFrontImageUrl(detail),
    visibility: detail.visibility,
    updatedAt: detail.updatedAt,
  }
}

function nowIso() {
  return new Date().toISOString()
}

export type CreateGalleryInput = {
  name: string
  description?: string
  visibility?: MediaVisibility
  frontImage?: File
}

export function useMockMediaStore() {
  const [, force] = useState(0)

  useEffect(() => {
    const listener = () => force((v) => v + 1)
    listeners.add(listener)
    return () => {
      listeners.delete(listener)
    }
  }, [])

  const createGallery = useCallback((input: CreateGalleryInput) => {
    const id = nextId++
    const trimmedDescription = input.description?.trim()
    const detail: GalleryDetail = {
      id,
      name: input.name.trim(),
      description: trimmedDescription || undefined,
      visibility: input.visibility,
      assetLimit: 20,
      updatedAt: nowIso(),
      assets: input.frontImage
        ? [
            {
              id: nextId++,
              fileName: input.frontImage.name,
              fileUrl: URL.createObjectURL(input.frontImage),
              mimeType: input.frontImage.type,
              fileSize: input.frontImage.size,
              uploadedAt: nowIso(),
            },
          ]
        : [],
    }
    galleries.set(id, detail)
    console.log('[mockMediaStore] createGallery', { input, detail })
    notify()
    return detail
  }, [])

  const getGallery = useCallback(
    (id: number): GalleryDetail | undefined => galleries.get(id),
    [],
  )

  const uploadAssets = useCallback(
    (galleryId: number, files: File[], altText: string) => {
      const gallery = galleries.get(galleryId)
      if (!gallery) {
        return
      }

      const newAssets: GalleryAsset[] = files.map((file) => ({
        id: nextId++,
        fileName: file.name,
        fileUrl: URL.createObjectURL(file),
        mimeType: file.type,
        fileSize: file.size,
        altText: altText.trim() || undefined,
        uploadedAt: nowIso(),
      }))

      const updated: GalleryDetail = {
        ...gallery,
        assets: [...(gallery.assets ?? []), ...newAssets],
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] uploadAssets', {
        galleryId,
        altText,
        newAssets,
      })
      notify()
    },
    [],
  )

  const deleteAsset = useCallback(
    (galleryId: number, asset: GalleryAsset) => {
      const gallery = galleries.get(galleryId)
      if (!gallery) {
        return
      }

      const updated: GalleryDetail = {
        ...gallery,
        assets: (gallery.assets ?? []).filter((item) => item.id !== asset.id),
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] deleteAsset', { galleryId, asset })
      notify()
    },
    [],
  )

  return {
    galleries: Array.from(galleries.values()).map(summarize),
    createGallery,
    getGallery,
    uploadAssets,
    deleteAsset,
  }
}
