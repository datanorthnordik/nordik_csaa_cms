import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import {
  CloseIcon,
  DeleteIcon,
  DownloadIcon,
  EditIcon,
  LinkIcon,
  SpecsIcon,
} from '../icons'
import {
  getGalleryAssetDetails,
  getGalleryAssetTitle,
} from '../../lib/galleryAssets'
import type {
  GalleryAsset,
  GalleryAssetContentPatch,
} from '../../types/media'
import styles from './SelectedImagePanel.module.css'

type SelectedImagePanelProps = {
  asset: GalleryAsset
  onClose?: () => void
  onUpdateAsset?: (asset: GalleryAsset, patch: GalleryAssetContentPatch) => void
  onDownload?: (asset: GalleryAsset) => void
  onDelete?: (asset: GalleryAsset) => void
  formatFileSize?: (bytes: number) => string
  formatUploadedAt?: (isoDate: string) => string
}

export function SelectedImagePanel({
  asset,
  onClose,
  onUpdateAsset,
  onDownload,
  onDelete,
  formatFileSize,
  formatUploadedAt,
}: SelectedImagePanelProps) {
  const { t } = useTranslation()
  const previewUrl = asset.fileUrl ?? asset.thumbnailUrl
  const assetTitle = getGalleryAssetTitle(asset)
  const assetDetails = getGalleryAssetDetails(asset)
  const assetLinkUrl = asset.linkUrl?.trim() ?? ''
  const [isEditing, setIsEditing] = useState(false)
  const [draftTitle, setDraftTitle] = useState(assetTitle)
  const [draftDetails, setDraftDetails] = useState(assetDetails)
  const [draftLinkUrl, setDraftLinkUrl] = useState(assetLinkUrl)

  useEffect(() => {
    setDraftTitle(assetTitle)
    setDraftDetails(assetDetails)
    setDraftLinkUrl(assetLinkUrl)
    setIsEditing(false)
  }, [asset.id, assetDetails, assetLinkUrl, assetTitle])

  const trimmedTitle = draftTitle.trim()
  const trimmedDetails = draftDetails.trim()
  const trimmedLinkUrl = draftLinkUrl.trim()
  const canSaveMetadata =
    Boolean(onUpdateAsset) &&
    trimmedTitle.length > 0 &&
    trimmedDetails.length > 0 &&
    (
      trimmedTitle !== assetTitle ||
      trimmedDetails !== assetDetails ||
      trimmedLinkUrl !== assetLinkUrl
    )

  const uploadedLine = buildUploadedLine(
    asset,
    formatUploadedAt,
    t('galleryManager.selected.uploadedBy', { name: asset.uploadedBy ?? '' }),
    t,
  )

  const hasAnyTechSpec = Boolean(
    asset.dimensions || asset.fileSize || asset.mimeType,
  )

  return (
    <aside className={styles.panel} aria-label={t('galleryManager.selected.title')}>
      <header className={styles.header}>
        <h2 className={styles.title}>{t('galleryManager.selected.title')}</h2>
        {onClose && (
          <button
            type="button"
            className={styles.closeButton}
            onClick={onClose}
            aria-label={t('galleryManager.selected.close')}
          >
            <CloseIcon />
          </button>
        )}
      </header>

      {previewUrl && (
        <div className={styles.preview}>
          <img
            src={previewUrl}
            alt={assetDetails || assetTitle}
            className={styles.previewImage}
          />
        </div>
      )}

      <div className={styles.heading}>
        <p className={styles.assetTitle}>{assetTitle}</p>
        <p className={styles.fileName}>{asset.fileName}</p>
        {uploadedLine && <p className={styles.uploadedLine}>{uploadedLine}</p>}
      </div>

      <section className={styles.section}>
        <div className={styles.sectionHeader}>
          <p className={styles.sectionLabel}>
            <EditIcon size={14} /> {t('galleryManager.selected.metadata.title')}
          </p>
          {onUpdateAsset && !isEditing && (
            <button
              type="button"
              className={styles.inlineActionButton}
              onClick={() => {
                setDraftTitle(assetTitle)
                setDraftDetails(assetDetails)
                setDraftLinkUrl(assetLinkUrl)
                setIsEditing(true)
              }}
            >
              {t('galleryManager.selected.metadata.edit')}
            </button>
          )}
        </div>

        {isEditing ? (
          <form
            className={styles.metaForm}
            onSubmit={(event) => {
              event.preventDefault()
              if (!onUpdateAsset || !canSaveMetadata) {
                return
              }

              onUpdateAsset(asset, {
                title: trimmedTitle,
                details: trimmedDetails,
                linkUrl: trimmedLinkUrl || undefined,
              })
              setIsEditing(false)
            }}
          >
            <label className={styles.metaField}>
              <span className={styles.metaLabel}>
                {t('galleryManager.selected.metadata.assetTitle')}
              </span>
              <input
                type="text"
                className={styles.metaInput}
                value={draftTitle}
                onChange={(event) => setDraftTitle(event.target.value)}
                aria-label={t('galleryManager.selected.metadata.assetTitle')}
              />
            </label>

            <label className={styles.metaField}>
              <span className={styles.metaLabel}>
                {t('galleryManager.selected.metadata.assetDetails')}
              </span>
              <textarea
                className={styles.metaTextarea}
                rows={4}
                value={draftDetails}
                onChange={(event) => setDraftDetails(event.target.value)}
                aria-label={t('galleryManager.selected.metadata.assetDetails')}
              />
            </label>

            <label className={styles.metaField}>
              <span className={styles.metaLabel}>
                {t('galleryManager.selected.metadata.linkUrl')}
              </span>
              <input
                type="text"
                className={styles.metaInput}
                value={draftLinkUrl}
                onChange={(event) => setDraftLinkUrl(event.target.value)}
                placeholder={t('galleryManager.selected.metadata.linkPlaceholder')}
                aria-label={t('galleryManager.selected.metadata.linkUrl')}
              />
            </label>

            <div className={styles.metaActions}>
              <button
                type="submit"
                className={styles.primaryActionButton}
                disabled={!canSaveMetadata}
              >
                {t('galleryManager.selected.metadata.save')}
              </button>
              <button
                type="button"
                className={styles.secondaryActionButton}
                onClick={() => {
                  setDraftTitle(assetTitle)
                  setDraftDetails(assetDetails)
                  setDraftLinkUrl(assetLinkUrl)
                  setIsEditing(false)
                }}
              >
                {t('galleryManager.selected.metadata.cancel')}
              </button>
            </div>
          </form>
        ) : (
          <div className={styles.metadataStack}>
            <div className={styles.metadataItem}>
              <p className={styles.metaLabel}>
                {t('galleryManager.selected.metadata.assetTitle')}
              </p>
              <p className={styles.metaValue}>{assetTitle}</p>
            </div>

            <div className={styles.metadataItem}>
              <p className={styles.metaLabel}>
                {t('galleryManager.selected.metadata.assetDetails')}
              </p>
              {assetDetails ? (
                <p className={styles.metaValue}>{assetDetails}</p>
              ) : (
                <p className={styles.metaEmpty}>
                  {t('galleryManager.selected.metadata.empty')}
                </p>
              )}
            </div>

            {assetLinkUrl && (
              <div className={styles.metadataItem}>
                <p className={styles.metaLabel}>
                  {t('galleryManager.selected.metadata.linkUrl')}
                </p>
                <a className={styles.usageLink} href={assetLinkUrl}>
                  {assetLinkUrl}
                </a>
              </div>
            )}
          </div>
        )}
      </section>

      {hasAnyTechSpec && (
        <section className={styles.section}>
          <p className={styles.sectionLabel}>
            <SpecsIcon /> {t('galleryManager.selected.specs.title')}
          </p>
          <dl className={styles.specs}>
            {asset.dimensions && (
              <>
                <dt>{t('galleryManager.selected.specs.dimensions')}</dt>
                <dd>
                  {t('galleryManager.selected.specs.dimensionsValue', {
                    width: asset.dimensions.width,
                    height: asset.dimensions.height,
                  })}
                </dd>
              </>
            )}
            {typeof asset.fileSize === 'number' && (
              <>
                <dt>{t('galleryManager.selected.specs.fileSize')}</dt>
                <dd>
                  {formatFileSize
                    ? formatFileSize(asset.fileSize)
                    : `${asset.fileSize} B`}
                </dd>
              </>
            )}
            {asset.mimeType && (
              <>
                <dt>{t('galleryManager.selected.specs.format')}</dt>
                <dd>{formatMimeType(asset.mimeType)}</dd>
              </>
            )}
          </dl>
        </section>
      )}

      {asset.usageTracking && asset.usageTracking.length > 0 && (
        <section className={styles.section}>
          <p className={styles.sectionLabel}>
            <LinkIcon /> {t('galleryManager.selected.usage.title')}
          </p>
          <ul className={styles.usageList}>
            {asset.usageTracking.map((usage, index) => (
              <li key={`${usage.path}-${index}`}>
                <a className={styles.usageLink} href={usage.path}>
                  {usage.label}
                </a>
              </li>
            ))}
          </ul>
        </section>
      )}

      <footer className={styles.footer}>
        {onDownload ? (
          <button
            type="button"
            className={styles.downloadButton}
            onClick={() => onDownload(asset)}
          >
            <DownloadIcon size={16} />
            {t('galleryManager.selected.actions.download')}
          </button>
        ) : (
          <a
            className={styles.downloadButton}
            href={asset.fileUrl}
            download={asset.fileName}
            rel="noopener"
          >
            <DownloadIcon size={16} />
            {t('galleryManager.selected.actions.download')}
          </a>
        )}
        <button
          type="button"
          className={styles.deleteButton}
          onClick={() => onDelete?.(asset)}
          aria-label={t('galleryManager.selected.actions.delete')}
        >
          <DeleteIcon size={16} />
        </button>
      </footer>
    </aside>
  )
}

function buildUploadedLine(
  asset: GalleryAsset,
  formatUploadedAt: ((iso: string) => string) | undefined,
  uploadedByText: string,
  t: ReturnType<typeof useTranslation>['t'],
) {
  const hasUploader = Boolean(asset.uploadedBy)
  const hasDate = Boolean(asset.uploadedAt)

  if (!hasUploader && !hasDate) {
    return null
  }

  if (hasUploader && hasDate && asset.uploadedAt) {
    return t('galleryManager.selected.uploadedLine', {
      name: asset.uploadedBy,
      date: formatUploadedAt ? formatUploadedAt(asset.uploadedAt) : asset.uploadedAt,
    })
  }

  if (hasUploader) {
    return uploadedByText
  }

  if (asset.uploadedAt) {
    return t('galleryManager.selected.uploadedAt', {
      date: formatUploadedAt ? formatUploadedAt(asset.uploadedAt) : asset.uploadedAt,
    })
  }

  return null
}

function formatMimeType(mime: string) {
  const parts = mime.split('/')
  return parts[1] ? parts[1].toUpperCase() : mime.toUpperCase()
}
