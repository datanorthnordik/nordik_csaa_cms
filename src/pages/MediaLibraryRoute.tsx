import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate } from 'react-router-dom'
import { fetchGalleryObjectUrl } from '../lib/galleryMedia'
import { fileToBase64 } from '../lib/fileUpload'
import { MediaLibraryPage } from './MediaLibraryPage'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import {
  createGallery,
  fetchMediaLibrary,
  selectMediaCreate,
  selectMediaLibrary,
} from '../store/mediaSlice'

export function MediaLibraryRoute() {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const listState = useAppSelector(selectMediaLibrary)
  const createState = useAppSelector(selectMediaCreate)
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
              data_base64: await fileToBase64(frontImage),
              alt_text: values.name.trim(),
            }
          : undefined,
      }

      const result = await dispatch(createGallery(payload)).unwrap()
      toast.success(result.message)
      navigate(`/media-library/${result.gallery.id}`)
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
      onCreate={handleCreate}
    />
  )
}
