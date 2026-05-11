import { useParams } from 'react-router-dom'
import { useMockMediaStore } from '../lib/mockMediaStore'
import { GalleryManagerPage } from './GalleryManagerPage'

export function GalleryManagerRoute() {
  const { galleryId } = useParams()
  const {
    getGallery,
    uploadAssets,
    deleteAsset,
    moveAsset,
    reorderAsset,
    setGalleryCover,
    saveGallery,
  } = useMockMediaStore()
  const numericId = galleryId ? Number.parseInt(galleryId, 10) : Number.NaN
  const gallery = Number.isFinite(numericId) ? getGallery(numericId) : undefined

  return (
    <GalleryManagerPage
      gallery={gallery}
      onUpdateGallery={
        gallery ? (patch) => saveGallery(gallery.id, patch) : undefined
      }
      onUploadAssets={
        gallery ? (uploads) => uploadAssets(gallery.id, uploads) : undefined
      }
      onDeleteAsset={
        gallery ? (asset) => deleteAsset(gallery.id, asset) : undefined
      }
      onMoveAsset={
        gallery
          ? (asset, delta) => moveAsset(gallery.id, asset, delta)
          : undefined
      }
      onReorderAsset={
        gallery
          ? (fromAsset, toAsset) =>
              reorderAsset(gallery.id, fromAsset, toAsset)
          : undefined
      }
      onSetCover={
        gallery ? (file) => setGalleryCover(gallery.id, file) : undefined
      }
      onDownloadAsset={(asset) => {
        console.log('[GalleryManagerRoute] download requested', asset)
      }}
      onSaveDraft={
        gallery
          ? () => saveGallery(gallery.id, { visibility: 'draft' })
          : undefined
      }
      onPublish={
        gallery
          ? () => saveGallery(gallery.id, { visibility: 'published' })
          : undefined
      }
    />
  )
}
