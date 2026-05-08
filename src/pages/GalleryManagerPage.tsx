import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { AddPhotoIcon, CloudUploadIcon, EditIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import { AssetThumbnail } from '../components/media/AssetThumbnail'
import { SelectedImagePanel } from '../components/media/SelectedImagePanel'
import { UploadDropzone } from '../components/media/UploadDropzone'
import type { GalleryAsset, GalleryDetail } from '../types/media'
import styles from '../styles/GalleryManagerPage.module.css'

type GalleryManagerPageProps = {
  gallery?: GalleryDetail
  loading?: boolean
  error?: string
  onPublish?: () => void
  onEditDescription?: () => void
  onUploadAssets?: (files: File[], altText: string) => void
  onDeleteAsset?: (asset: GalleryAsset) => void
  onDownloadAsset?: (asset: GalleryAsset) => void
  formatRelativeTime?: (isoDate: string) => string
  formatFileSize?: (bytes: number) => string
  formatUploadedAt?: (isoDate: string) => string
}

export function GalleryManagerPage({
  gallery,
  loading = false,
  error,
  onPublish,
  onEditDescription,
  onUploadAssets,
  onDeleteAsset,
  onDownloadAsset,
  formatRelativeTime,
  formatFileSize,
  formatUploadedAt,
}: GalleryManagerPageProps = {}) {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const { galleryId } = useParams()
  const [selectedAssetId, setSelectedAssetId] = useState<number | null>(null)
  const [pendingFiles, setPendingFiles] = useState<File[]>([])
  const [altText, setAltText] = useState('')

  const selectedAsset =
    gallery?.assets && selectedAssetId !== null
      ? gallery.assets.find((asset) => asset.id === selectedAssetId) ?? null
      : null

  const assetCount = gallery?.assets?.length ?? 0
  const showUploadTile =
    typeof gallery?.assetLimit === 'number' && assetCount < gallery.assetLimit

  function handleUpload(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault()
    if (!pendingFiles.length) {
      return
    }

    onUploadAssets?.(pendingFiles, altText.trim())
    setPendingFiles([])
    setAltText('')
  }

  return (
    <CmsAppShell activeKey="media">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('mediaLibrary.breadcrumb.media') },
            {
              label: t('mediaLibrary.breadcrumb.galleryLibrary'),
              to: '/media-library',
            },
            { label: gallery?.name ?? t('galleryManager.breadcrumb.gallery') },
          ]}
        />

        {loading && !gallery ? (
          <div className={styles.loaderWrap}>
            <Loader label={t('galleryManager.loading')} />
          </div>
        ) : !gallery ? (
          <div className={styles.emptyState}>
            <p className={styles.emptyTitle}>
              {t('galleryManager.notFound.title')}
            </p>
            <p className={styles.emptyText}>
              {t('galleryManager.notFound.text', { id: galleryId ?? '' })}
            </p>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => navigate('/media-library')}
            >
              {t('galleryManager.notFound.back')}
            </button>
          </div>
        ) : (
          <div className={styles.layout}>
            <div className={styles.main}>
              <header className={styles.galleryHeader}>
                <h1 className={styles.title}>{gallery.name}</h1>
                <div className={styles.metaRow}>
                  {typeof gallery.assetLimit === 'number' && (
                    <span className={styles.metaItem}>
                      {t('galleryManager.header.assetsOf', {
                        count: assetCount,
                        limit: gallery.assetLimit,
                      })}
                    </span>
                  )}
                  {gallery.updatedAt && (
                    <>
                      {typeof gallery.assetLimit === 'number' && (
                        <span className={styles.metaDot} aria-hidden="true">
                          •
                        </span>
                      )}
                      <span className={styles.metaItem}>
                        {formatRelativeTime
                          ? t('galleryManager.header.updatedRelative', {
                              relative: formatRelativeTime(gallery.updatedAt),
                            })
                          : t('galleryManager.header.updatedAt', {
                              date: gallery.updatedAt,
                            })}
                      </span>
                    </>
                  )}
                </div>

                {gallery.description && (
                  <p className={styles.description}>{gallery.description}</p>
                )}

                <div className={styles.galleryActions}>
                  {onEditDescription && (
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={onEditDescription}
                    >
                      <EditIcon />
                      {t('galleryManager.header.editDescription')}
                    </button>
                  )}
                  <button
                    type="button"
                    className={styles.publishButton}
                    onClick={onPublish}
                    disabled={!onPublish}
                  >
                    {t('galleryManager.header.publish')}
                  </button>
                </div>
              </header>

              {error && <p className={styles.errorText}>{error}</p>}

              {gallery.assets === undefined ? (
                <div className={styles.loaderWrap}>
                  <Loader label={t('galleryManager.assetsLoading')} />
                </div>
              ) : (
                <div className={styles.assetGrid}>
                  {gallery.assets.map((asset) => (
                    <AssetThumbnail
                      key={asset.id}
                      asset={asset}
                      selected={asset.id === selectedAssetId}
                      onSelect={(item) => setSelectedAssetId(item.id)}
                    />
                  ))}
                  {showUploadTile && gallery.assetLimit !== undefined && (
                    <div className={styles.uploadTile} aria-hidden="true">
                      <AddPhotoIcon size={28} />
                      <span className={styles.uploadTileLabel}>
                        {t('galleryManager.uploads.maxLabel', {
                          count: gallery.assetLimit,
                        })}
                      </span>
                    </div>
                  )}
                </div>
              )}

              <section className={styles.uploadSection}>
                <h2 className={styles.uploadTitle}>
                  {t('galleryManager.uploads.title')}
                </h2>
                <form className={styles.uploadForm} onSubmit={handleUpload}>
                  <UploadDropzone
                    icon={<CloudUploadIcon />}
                    accept="image/png,image/jpeg"
                    multiple
                    label={t('galleryManager.uploads.dropLabel')}
                    hint={t('galleryManager.uploads.dropHint')}
                    onFiles={(files) => setPendingFiles(files)}
                  />
                  <div className={styles.uploadDetails}>
                    <label className={styles.field}>
                      <span className={styles.fieldLabel}>
                        {t('galleryManager.uploads.altLabel')}
                        <span className={styles.required} aria-hidden="true">
                          *
                        </span>
                      </span>
                      <textarea
                        className={styles.textarea}
                        rows={3}
                        value={altText}
                        onChange={(event) => setAltText(event.target.value)}
                        placeholder={t('galleryManager.uploads.altPlaceholder')}
                      />
                    </label>
                    {pendingFiles.length > 0 && (
                      <p className={styles.pendingFiles}>
                        {t('galleryManager.uploads.pendingFiles', {
                          count: pendingFiles.length,
                        })}
                      </p>
                    )}
                    <div className={styles.uploadActions}>
                      <button
                        type="submit"
                        className={styles.primaryButton}
                        disabled={!pendingFiles.length || !altText.trim()}
                      >
                        {t('galleryManager.uploads.submit')}
                      </button>
                    </div>
                  </div>
                </form>
              </section>
            </div>

            {selectedAsset && (
              <SelectedImagePanel
                asset={selectedAsset}
                onClose={() => setSelectedAssetId(null)}
                onDownload={onDownloadAsset}
                onDelete={onDeleteAsset}
                formatFileSize={formatFileSize}
                formatUploadedAt={formatUploadedAt}
              />
            )}
          </div>
        )}
      </div>
    </CmsAppShell>
  )
}
