import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate } from 'react-router-dom'
import { fetchGalleryObjectUrl } from '../lib/galleryMedia'
import { MediaLibraryPage } from './MediaLibraryPage'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import {
  createGallery,
  deleteGallery,
  fetchMediaLibrary,
  selectMediaCreate,
  selectMediaLibrary,
} from '../store/mediaSlice'
import type { GallerySummary } from '../types/media'

export function MediaLibraryRoute() {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const listState = useAppSelector(selectMediaLibrary)
  const createState = useAppSelector(selectMediaCreate)
  const deleteState = useAppSelector((state) => state.media.deleteGallery)
  const [previewUrls, setPreviewUrls] = useState<Record<number, string>>({})

  useEffect(() => {
    void dispatch(fetchMediaLibrary())
  }, [dispatch])

  useEffect(() => {
    if (!listState.items.length) {
      setPreviewUrls((current) => {
        Object.values(current).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return {}
      })
      return
    }

    let cancelled = false

    async function loadPreviews() {
      const loadedEntries = await Promise.all(
        listState.items.map(async (gallery) => {
          if (!gallery.frontImagePath) {
            return null
          }

          try {
            const objectUrl = await fetchGalleryObjectUrl(gallery.frontImagePath)
            return [gallery.id, objectUrl] as const
          } catch {
            return null
          }
        }),
      )

      const nextUrls = loadedEntries.reduce<Record<number, string>>((result, entry) => {
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
  }, [listState.items])

  const galleries = useMemo(
    () =>
      listState.items.map((gallery) => ({
        ...gallery,
        frontImageUrl: previewUrls[gallery.id] ?? undefined,
      })),
    [listState.items, previewUrls],
  )

  async function handleCreate(
    values: { name: string; description: string; visibility: 'draft' | 'published' },
    frontImage?: File,
  ) {
    try {
      const payload = {
        name: values.name.trim(),
        description: values.description.trim() || undefined,
        published: values.visibility === 'published',
        cover_image: frontImage
          ? {
              file_name: frontImage.name,
              mime_type: frontImage.type || 'application/octet-stream',
              alt_text: values.name.trim(),
            }
          : undefined,
        ...(frontImage
          ? {
              coverImageFile: frontImage,
            }
          : {}),
      }

      const result = await dispatch(createGallery(payload)).unwrap()
      toast.success(result.message)
      navigate(`/media-library/${result.gallery.id}`)
      return true
    } catch {
      return false
    }
  }

  async function handleDelete(gallery: GallerySummary) {
    try {
      const result = await dispatch(deleteGallery(gallery.id)).unwrap()
      toast.success(result.result.message)
      return true
    } catch {
      return false
    }
  }

  return (
    <MediaLibraryPage
      galleries={galleries}
      loading={listState.status === 'loading'}
      error={listState.error ?? undefined}
      creating={createState.status === 'loading'}
      deleting={deleteState.status === 'loading'}
      onCreate={handleCreate}
      onDelete={handleDelete}
    />
  )
}
