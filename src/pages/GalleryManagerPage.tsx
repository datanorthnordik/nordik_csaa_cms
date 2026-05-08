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
  saving?: boolean
  error?: string
  onSaveDraft?: () => void
  onSaveChanges?: () => void
  onPublish?: () => void
  onEditDescription?: () => void
  onUploadAssets?: (files: File[], altText: string) => void
  onDeleteAsset?: (asset: GalleryAsset) => void
  onDownloadAsset?: (asset: GalleryAsset) => void
  onMoveAsset?: (asset: GalleryAsset, delta: number) => void
  formatRelativeTime?: (isoDate: string) => string
  formatFileSize?: (bytes: number) => string
  formatUploadedAt?: (isoDate: string) => string
}

export function GalleryManagerPage({
  gallery,
  loading = false,
  saving = false,
  error,
  onSaveDraft,
  onSaveChanges,
  onPublish,
  onEditDescription,
  onUploadAssets,
  onDeleteAsset,
  onDownloadAsset,
  onMoveAsset,
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

                {onEditDescription && (
                  <div className={styles.galleryActions}>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={onEditDescription}
                    >
                      <EditIcon />
                      {t('galleryManager.header.editDescription')}
                    </button>
                  </div>
                )}
              </header>

              {error && <p className={styles.errorText}>{error}</p>}

              {gallery.assets === undefined ? (
                <div className={styles.loaderWrap}>
                  <Loader label={t('galleryManager.assetsLoading')} />
                </div>
              ) : (
                <div className={styles.assetGrid}>
                  {gallery.assets.map((asset, index) => (
                    <AssetThumbnail
                      key={asset.id}
                      asset={asset}
                      selected={asset.id === selectedAssetId}
                      onSelect={(item) => setSelectedAssetId(item.id)}
                      onMoveLeft={
                        onMoveAsset && index > 0
                          ? (item) => onMoveAsset(item, -1)
                          : undefined
                      }
                      onMoveRight={
                        onMoveAsset &&
                        gallery.assets &&
                        index < gallery.assets.length - 1
                          ? (item) => onMoveAsset(item, 1)
                          : undefined
                      }
                      onDelete={
                        onDeleteAsset
                          ? (item) => {
                              if (item.id === selectedAssetId) {
                                setSelectedAssetId(null)
                              }
                              onDeleteAsset(item)
                            }
                          : undefined
                      }
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

              <section className={styles.actionsCard}>
                <div className={styles.actionsHeader}>
                  <h2 className={styles.actionsTitle}>
                    {t('galleryManager.actions.title')}
                  </h2>
                  <p className={styles.actionsHint}>
                    {t('galleryManager.actions.hint')}
                  </p>
                </div>

                <div className={styles.statusChipRow}>
                  <span
                    className={[
                      styles.statusChip,
                      gallery.visibility === 'published'
                        ? styles.statusPublished
                        : styles.statusDraft,
                    ].join(' ')}
                  >
                    {gallery.visibility === 'published'
                      ? t('galleryManager.actions.statusPublished')
                      : t('galleryManager.actions.statusDraft')}
                  </span>
                </div>

                <div className={styles.actionStack}>
                  <button
                    type="button"
                    className={styles.secondaryButton}
                    onClick={onSaveDraft}
                    disabled={!onSaveDraft || saving}
                  >
                    {t('galleryManager.actions.saveDraft')}
                  </button>
                  <button
                    type="button"
                    className={styles.primaryButton}
                    onClick={onSaveChanges}
                    disabled={!onSaveChanges || saving}
                  >
                    {saving
                      ? t('galleryManager.actions.saving')
                      : t('galleryManager.actions.saveChanges')}
                  </button>
                  {gallery.visibility !== 'published' && onPublish && (
                    <button
                      type="button"
                      className={styles.publishButton}
                      onClick={onPublish}
                      disabled={saving}
                    >
                      {t('galleryManager.actions.publish')}
                    </button>
                  )}
                </div>
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
