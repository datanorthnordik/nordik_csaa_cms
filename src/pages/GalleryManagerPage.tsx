import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { CloseIcon, CloudUploadIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import { AssetThumbnail } from '../components/media/AssetThumbnail'
import { GalleryStatusChip } from '../components/media/GalleryStatusChip'
import { SelectedImagePanel } from '../components/media/SelectedImagePanel'
import { UploadDropzone } from '../components/media/UploadDropzone'
import {
  getGalleryAssetDetails,
  getGalleryAssetTitle,
  suggestGalleryAssetTitle,
} from '../lib/galleryAssets'
import type {
  GalleryAsset,
  GalleryAssetContentPatch,
  GalleryDetail,
} from '../types/media'
import styles from '../styles/GalleryManagerPage.module.css'

export type GalleryUpdatePatch = Partial<Pick<GalleryDetail, 'name' | 'description'>>

export type PendingUploadInput = {
  file: File
  title: string
  details: string
}

type GalleryManagerPageProps = {
  gallery?: GalleryDetail
  loading?: boolean
  saving?: boolean
  error?: string
  onUpdateGallery?: (patch: GalleryUpdatePatch) => void
  onSaveDraft?: () => void
  onPublish?: () => void
  onUploadAssets?: (uploads: PendingUploadInput[]) => void
  onUpdateAsset?: (asset: GalleryAsset, patch: GalleryAssetContentPatch) => void
  onDeleteAsset?: (asset: GalleryAsset) => void
  onDownloadAsset?: (asset: GalleryAsset) => void
  onMoveAsset?: (asset: GalleryAsset, delta: number) => void
  onReorderAsset?: (fromAsset: GalleryAsset, toAsset: GalleryAsset) => void
  onSetCover?: (file: File | null) => void
  formatRelativeTime?: (isoDate: string) => string
  formatFileSize?: (bytes: number) => string
  formatUploadedAt?: (isoDate: string) => string
}

type PendingUpload = {
  id: string
  file: File
  previewUrl: string
  title: string
  details: string
}

function makePendingId() {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

export function GalleryManagerPage({
  gallery,
  loading = false,
  saving = false,
  error,
  onUpdateGallery,
  onSaveDraft,
  onPublish,
  onUploadAssets,
  onUpdateAsset,
  onDeleteAsset,
  onDownloadAsset,
  onMoveAsset,
  onReorderAsset,
  onSetCover,
  formatRelativeTime,
  formatFileSize,
  formatUploadedAt,
}: GalleryManagerPageProps = {}) {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const { galleryId } = useParams()
  const [selectedAssetId, setSelectedAssetId] = useState<number | null>(null)
  const [pendingUploads, setPendingUploads] = useState<PendingUpload[]>([])
  const [draggedAssetId, setDraggedAssetId] = useState<number | null>(null)
  const [dragOverAssetId, setDragOverAssetId] = useState<number | null>(null)

  const selectedAsset =
    gallery?.assets && selectedAssetId !== null
      ? gallery.assets.find((asset) => asset.id === selectedAssetId) ?? null
      : null

  const assetCount = gallery?.assets?.length ?? 0
  const assetLimit = gallery?.assetLimit
  const remainingSlots =
    typeof assetLimit === 'number'
      ? Math.max(0, assetLimit - assetCount - pendingUploads.length)
      : Infinity
  const isLimitReached =
    typeof assetLimit === 'number' && assetCount >= assetLimit
  const willExceedLimit =
    typeof assetLimit === 'number' &&
    assetCount + pendingUploads.length > assetLimit
  const allPendingMetadataFilled =
    pendingUploads.length > 0 &&
    pendingUploads.every(
      (upload) =>
        upload.title.trim().length > 0 && upload.details.trim().length > 0,
    )

  function handleFilesDropped(files: File[]) {
    const accepted =
      remainingSlots === Infinity ? files : files.slice(0, remainingSlots)
    const rejected = files.length - accepted.length

    if (rejected > 0) {
      console.warn(
        `[GalleryManagerPage] Rejected ${rejected} file(s) — gallery limit reached`,
      )
    }

    if (!accepted.length) {
      return
    }

    const additions: PendingUpload[] = accepted.map((file) => ({
      id: makePendingId(),
      file,
      previewUrl: URL.createObjectURL(file),
      title: suggestGalleryAssetTitle(file.name),
      details: '',
    }))

    setPendingUploads((current) => [...current, ...additions])
  }

  function updatePendingUpload(
    id: string,
    field: keyof Pick<PendingUpload, 'title' | 'details'>,
    value: string,
  ) {
    setPendingUploads((current) =>
      current.map((upload) =>
        upload.id === id ? { ...upload, [field]: value } : upload,
      ),
    )
  }

  function removePending(id: string) {
    setPendingUploads((current) => {
      const target = current.find((upload) => upload.id === id)
      if (target) {
        URL.revokeObjectURL(target.previewUrl)
      }
      return current.filter((upload) => upload.id !== id)
    })
  }

  function handleUpload(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault()
    if (!pendingUploads.length || !allPendingMetadataFilled || willExceedLimit) {
      return
    }

    onUploadAssets?.(
      pendingUploads.map((upload) => ({
        file: upload.file,
        title: upload.title.trim(),
        details: upload.details.trim(),
      })),
    )

    pendingUploads.forEach((upload) => URL.revokeObjectURL(upload.previewUrl))
    setPendingUploads([])
  }

  function commitName(value: string) {
    if (!onUpdateGallery || !gallery) {
      return
    }
    const trimmed = value.trim()
    if (!trimmed || trimmed === gallery.name) {
      return
    }
    onUpdateGallery({ name: trimmed })
  }

  function commitDescription(value: string) {
    if (!onUpdateGallery || !gallery) {
      return
    }
    const trimmed = value.trim()
    const current = gallery.description ?? ''
    if (trimmed === current) {
      return
    }
    onUpdateGallery({ description: trimmed || undefined })
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
                <div className={styles.headerTop}>
                  <GalleryStatusChip visibility={gallery.visibility} />
                </div>

                <h1 className={styles.titleRow}>
                  <input
                    key={`name-${gallery.id}`}
                    type="text"
                    className={styles.titleInput}
                    defaultValue={gallery.name}
                    placeholder={t('galleryManager.header.namePlaceholder')}
                    aria-label={t('galleryManager.header.namePlaceholder')}
                    disabled={!onUpdateGallery}
                    onBlur={(event) => commitName(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter') {
                        event.currentTarget.blur()
                      }
                    }}
                  />
                </h1>

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
              </header>

              <section className={styles.coverCard}>
                <div className={styles.coverCardHeader}>
                  <h2 className={styles.coverCardTitle}>
                    {t('galleryManager.cover.title')}
                  </h2>
                  <p className={styles.coverCardHint}>
                    {t('galleryManager.cover.hint')}
                  </p>
                </div>

                <div className={styles.coverCardBody}>
                  <div className={styles.coverArea}>
                    {gallery.coverImage ? (
                      <>
                        <img
                          src={gallery.coverImage.fileUrl}
                          alt={
                            getGalleryAssetDetails(gallery.coverImage) ||
                            getGalleryAssetTitle(gallery.coverImage)
                          }
                          className={styles.coverImage}
                        />
                        <div className={styles.coverActions}>
                          <label className={styles.coverReplaceButton}>
                            <input
                              type="file"
                              accept="image/png,image/jpeg"
                              hidden
                              disabled={!onSetCover}
                              onChange={(event) => {
                                const file = event.target.files?.[0]
                                if (file) {
                                  onSetCover?.(file)
                                }
                                event.target.value = ''
                              }}
                            />
                            {t('galleryManager.cover.replace')}
                          </label>
                          <button
                            type="button"
                            className={styles.coverRemoveButton}
                            onClick={() => onSetCover?.(null)}
                            disabled={!onSetCover}
                          >
                            {t('galleryManager.cover.remove')}
                          </button>
                        </div>
                      </>
                    ) : (
                      <UploadDropzone
                        icon={<CloudUploadIcon />}
                        accept="image/png,image/jpeg"
                        label={t('galleryManager.cover.dropLabel')}
                        hint={t('galleryManager.cover.dropHint')}
                        disabled={!onSetCover}
                        onFiles={(files) => {
                          const file = files[0]
                          if (file) {
                            onSetCover?.(file)
                          }
                        }}
                      />
                    )}
                  </div>

                  <label className={styles.coverDescription}>
                    <span className={styles.fieldLabel}>
                      {t('galleryManager.cover.descriptionLabel')}
                    </span>
                    <textarea
                      key={`desc-${gallery.id}`}
                      className={styles.descriptionInput}
                      defaultValue={gallery.description ?? ''}
                      placeholder={t(
                        'galleryManager.header.descriptionPlaceholder',
                      )}
                      aria-label={t(
                        'galleryManager.header.descriptionPlaceholder',
                      )}
                      rows={6}
                      disabled={!onUpdateGallery}
                      onBlur={(event) => commitDescription(event.target.value)}
                    />
                  </label>
                </div>
              </section>

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
                      draggable={Boolean(onReorderAsset)}
                      isDragging={asset.id === draggedAssetId}
                      isDragOver={
                        asset.id === dragOverAssetId &&
                        asset.id !== draggedAssetId
                      }
                      onDragStart={(item) => setDraggedAssetId(item.id)}
                      onDragEnd={() => {
                        setDraggedAssetId(null)
                        setDragOverAssetId(null)
                      }}
                      onDragEnter={(item) => setDragOverAssetId(item.id)}
                      onDragLeave={(item) => {
                        setDragOverAssetId((current) =>
                          current === item.id ? null : current,
                        )
                      }}
                      onDragDrop={(target) => {
                        if (
                          !onReorderAsset ||
                          draggedAssetId === null ||
                          draggedAssetId === target.id ||
                          !gallery.assets
                        ) {
                          return
                        }
                        const fromAsset = gallery.assets.find(
                          (item) => item.id === draggedAssetId,
                        )
                        if (fromAsset) {
                          onReorderAsset(fromAsset, target)
                        }
                        setDraggedAssetId(null)
                        setDragOverAssetId(null)
                      }}
                    />
                  ))}
                </div>
              )}

              <section className={styles.uploadSection}>
                <div className={styles.uploadHeader}>
                  <h2 className={styles.uploadTitle}>
                    {t('galleryManager.uploads.title')}
                  </h2>
                  {typeof assetLimit === 'number' && (
                    <p className={styles.uploadNote}>
                      {t('galleryManager.uploads.maxNote', {
                        count: assetLimit,
                      })}
                    </p>
                  )}
                </div>
                <form className={styles.uploadForm} onSubmit={handleUpload}>
                  <UploadDropzone
                    icon={<CloudUploadIcon />}
                    accept="image/png,image/jpeg"
                    multiple
                    label={t('galleryManager.uploads.dropLabel')}
                    hint={t('galleryManager.uploads.dropHint')}
                    disabled={isLimitReached}
                    onFiles={handleFilesDropped}
                  />

                  {isLimitReached && (
                    <p className={styles.limitWarning} role="alert">
                      {t('galleryManager.uploads.limitReached', {
                        limit: assetLimit,
                      })}
                    </p>
                  )}

                  {pendingUploads.length > 0 && (
                    <>
                      <ul className={styles.pendingList}>
                        {pendingUploads.map((upload) => (
                          <li key={upload.id} className={styles.pendingItem}>
                            <img
                              src={upload.previewUrl}
                              alt=""
                              className={styles.pendingThumb}
                            />
                            <div className={styles.pendingFields}>
                              <p className={styles.pendingFileName}>
                                {upload.file.name}
                              </p>
                              <label className={styles.pendingAltLabel}>
                                <span className={styles.fieldLabel}>
                                  {t('galleryManager.uploads.titleLabel')}
                                  <span
                                    className={styles.required}
                                    aria-hidden="true"
                                  >
                                    *
                                  </span>
                                </span>
                                <input
                                  type="text"
                                  className={styles.pendingTextInput}
                                  value={upload.title}
                                  onChange={(event) =>
                                    updatePendingUpload(
                                      upload.id,
                                      'title',
                                      event.target.value,
                                    )
                                  }
                                  placeholder={t(
                                    'galleryManager.uploads.titlePlaceholder',
                                  )}
                                  aria-label={t(
                                    'galleryManager.uploads.titleLabelFor',
                                    { name: upload.file.name },
                                  )}
                                />
                              </label>
                              <label className={styles.pendingAltLabel}>
                                <span className={styles.fieldLabel}>
                                  {t('galleryManager.uploads.detailsLabel')}
                                  <span
                                    className={styles.required}
                                    aria-hidden="true"
                                  >
                                    *
                                  </span>
                                </span>
                                <textarea
                                  className={styles.pendingAltInput}
                                  rows={2}
                                  value={upload.details}
                                  onChange={(event) =>
                                    updatePendingUpload(
                                      upload.id,
                                      'details',
                                      event.target.value,
                                    )
                                  }
                                  placeholder={t(
                                    'galleryManager.uploads.detailsPlaceholder',
                                  )}
                                  aria-label={t(
                                    'galleryManager.uploads.detailsLabelFor',
                                    { name: upload.file.name },
                                  )}
                                />
                              </label>
                            </div>
                            <button
                              type="button"
                              className={styles.pendingRemove}
                              onClick={() => removePending(upload.id)}
                              aria-label={t(
                                'galleryManager.uploads.removeUpload',
                                { name: upload.file.name },
                              )}
                            >
                              <CloseIcon size={16} />
                            </button>
                          </li>
                        ))}
                      </ul>

                      {willExceedLimit && (
                        <p className={styles.limitWarning} role="alert">
                          {t('galleryManager.uploads.limitExceeded', {
                            count: assetCount + pendingUploads.length,
                            limit: assetLimit,
                          })}
                        </p>
                      )}
                    </>
                  )}

                  <div className={styles.uploadActions}>
                    <button
                      type="submit"
                      className={styles.primaryButton}
                      disabled={
                        !pendingUploads.length ||
                        !allPendingMetadataFilled ||
                        willExceedLimit
                      }
                    >
                      {t('galleryManager.uploads.submit')}
                    </button>
                  </div>
                </form>
              </section>

              <div className={styles.actionsCard}>
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
                    onClick={onPublish}
                    disabled={!onPublish || saving}
                  >
                    {t('galleryManager.actions.publish')}
                  </button>
                </div>
              </div>
            </div>

            {selectedAsset && (
              <SelectedImagePanel
                asset={selectedAsset}
                onClose={() => setSelectedAssetId(null)}
                onUpdateAsset={onUpdateAsset}
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
