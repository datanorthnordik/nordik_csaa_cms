import { useCallback, useEffect, useState } from 'react'
import { suggestGalleryAssetTitle } from './galleryAssets'
import type {
  GalleryAsset,
  GalleryAssetContentPatch,
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
  return detail.coverImage?.fileUrl ?? detail.assets?.[0]?.fileUrl
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

function assetFromFile(file: File, id: number): GalleryAsset {
  return {
    id,
    fileName: file.name,
    fileUrl: URL.createObjectURL(file),
    mimeType: file.type,
    fileSize: file.size,
    uploadedAt: nowIso(),
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
      coverImage: input.frontImage
        ? assetFromFile(input.frontImage, nextId++)
        : undefined,
      assets: [],
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
    (
      galleryId: number,
      uploads: { file: File; title: string; details: string }[],
    ) => {
      const gallery = galleries.get(galleryId)
      if (!gallery) {
        return
      }

      const limit = gallery.assetLimit ?? Number.POSITIVE_INFINITY
      const currentCount = gallery.assets?.length ?? 0
      const acceptable = Math.max(0, limit - currentCount)
      const accepted = uploads.slice(0, acceptable)

      if (accepted.length < uploads.length) {
        console.warn(
          '[mockMediaStore] uploadAssets rejected',
          uploads.length - accepted.length,
          'file(s) — gallery limit reached',
        )
      }

      if (!accepted.length) {
        return
      }

      const newAssets: GalleryAsset[] = accepted.map(({ file, title, details }) => ({
        id: nextId++,
        fileName: file.name,
        fileUrl: URL.createObjectURL(file),
        mimeType: file.type,
        fileSize: file.size,
        title: title.trim() || suggestGalleryAssetTitle(file.name),
        details: details.trim() || undefined,
        altText: details.trim() || undefined,
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
        uploads,
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

  const updateAsset = useCallback(
    (galleryId: number, assetId: number, patch: GalleryAssetContentPatch) => {
      const gallery = galleries.get(galleryId)
      if (!gallery?.assets) {
        return
      }

      const nextAssets = gallery.assets.map((asset) => {
        if (asset.id !== assetId) {
          return asset
        }

        const title =
          patch.title.trim() || suggestGalleryAssetTitle(asset.fileName)
        const details = patch.details.trim()

        return {
          ...asset,
          title,
          details: details || undefined,
          altText: details || undefined,
        }
      })

      const updated: GalleryDetail = {
        ...gallery,
        assets: nextAssets,
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] updateAsset', { galleryId, assetId, patch })
      notify()
    },
    [],
  )

  const moveAsset = useCallback(
    (galleryId: number, asset: GalleryAsset, delta: number) => {
      const gallery = galleries.get(galleryId)
      if (!gallery?.assets) {
        return
      }

      const fromIndex = gallery.assets.findIndex((item) => item.id === asset.id)
      const toIndex = fromIndex + delta
      if (
        fromIndex === -1 ||
        toIndex < 0 ||
        toIndex >= gallery.assets.length
      ) {
        return
      }

      const next = [...gallery.assets]
      const [moved] = next.splice(fromIndex, 1)
      next.splice(toIndex, 0, moved)

      const updated: GalleryDetail = {
        ...gallery,
        assets: next,
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] moveAsset', {
        galleryId,
        assetId: asset.id,
        fromIndex,
        toIndex,
      })
      notify()
    },
    [],
  )

  const reorderAsset = useCallback(
    (galleryId: number, fromAsset: GalleryAsset, toAsset: GalleryAsset) => {
      const gallery = galleries.get(galleryId)
      if (!gallery?.assets) {
        return
      }

      const fromIndex = gallery.assets.findIndex(
        (item) => item.id === fromAsset.id,
      )
      const toIndex = gallery.assets.findIndex(
        (item) => item.id === toAsset.id,
      )
      if (fromIndex === -1 || toIndex === -1 || fromIndex === toIndex) {
        return
      }

      const next = [...gallery.assets]
      const [moved] = next.splice(fromIndex, 1)
      next.splice(toIndex, 0, moved)

      const updated: GalleryDetail = {
        ...gallery,
        assets: next,
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] reorderAsset', {
        galleryId,
        fromId: fromAsset.id,
        toId: toAsset.id,
        fromIndex,
        toIndex,
      })
      notify()
    },
    [],
  )

  const setGalleryCover = useCallback(
    (galleryId: number, file: File | null) => {
      const gallery = galleries.get(galleryId)
      if (!gallery) {
        return
      }

      if (gallery.coverImage) {
        URL.revokeObjectURL(gallery.coverImage.fileUrl)
      }

      const nextCover = file ? assetFromFile(file, nextId++) : undefined

      const updated: GalleryDetail = {
        ...gallery,
        coverImage: nextCover,
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] setGalleryCover', {
        galleryId,
        fileName: file?.name ?? null,
      })
      notify()
    },
    [],
  )

  const saveGallery = useCallback(
    (
      galleryId: number,
      patch: Partial<
        Pick<GalleryDetail, 'visibility' | 'name' | 'description'>
      >,
    ) => {
      const gallery = galleries.get(galleryId)
      if (!gallery) {
        return
      }

      const updated: GalleryDetail = {
        ...gallery,
        ...patch,
        updatedAt: nowIso(),
      }

      galleries.set(galleryId, updated)
      console.log('[mockMediaStore] saveGallery', { galleryId, patch, updated })
      notify()
    },
    [],
  )

  return {
    galleries: Array.from(galleries.values()).map(summarize),
    createGallery,
    getGallery,
    uploadAssets,
    updateAsset,
    deleteAsset,
    moveAsset,
    reorderAsset,
    setGalleryCover,
    saveGallery,
  }
}
