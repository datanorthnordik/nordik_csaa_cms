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
    saveGallery,
  } = useMockMediaStore()
  const numericId = galleryId ? Number.parseInt(galleryId, 10) : Number.NaN
  const gallery = Number.isFinite(numericId) ? getGallery(numericId) : undefined

  return (
    <GalleryManagerPage
      gallery={gallery}
      onUploadAssets={
        gallery
          ? (files, altText) => uploadAssets(gallery.id, files, altText)
          : undefined
      }
      onDeleteAsset={
        gallery ? (asset) => deleteAsset(gallery.id, asset) : undefined
      }
      onMoveAsset={
        gallery
          ? (asset, delta) => moveAsset(gallery.id, asset, delta)
          : undefined
      }
      onDownloadAsset={(asset) => {
        console.log('[GalleryManagerRoute] download requested', asset)
      }}
      onSaveDraft={
        gallery
          ? () => saveGallery(gallery.id, { visibility: 'draft' })
          : undefined
      }
      onSaveChanges={
        gallery
          ? () => saveGallery(gallery.id, {})
          : undefined
      }
      onPublish={
        gallery
          ? () => saveGallery(gallery.id, { visibility: 'published' })
          : undefined
      }
      onEditDescription={() => {
        console.log('[GalleryManagerRoute] edit description clicked', gallery)
      }}
    />
  )
}
