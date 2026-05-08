import { useParams } from 'react-router-dom'
import { useMockMediaStore } from '../lib/mockMediaStore'
import { GalleryManagerPage } from './GalleryManagerPage'

export function GalleryManagerRoute() {
  const { galleryId } = useParams()
  const { getGallery, uploadAssets, deleteAsset } = useMockMediaStore()
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
      onDownloadAsset={(asset) => {
        console.log('[GalleryManagerRoute] download requested', asset)
      }}
      onPublish={() => {
        console.log('[GalleryManagerRoute] publish clicked', gallery)
      }}
      onEditDescription={() => {
        console.log('[GalleryManagerRoute] edit description clicked', gallery)
      }}
    />
  )
}
