import { useTranslation } from 'react-i18next'
import {
  CheckCircleIcon,
  CloseIcon,
  DeleteIcon,
  DownloadIcon,
  LinkIcon,
  SpecsIcon,
} from '../icons'
import type { GalleryAsset } from '../../types/media'
import styles from './SelectedImagePanel.module.css'

type SelectedImagePanelProps = {
  asset: GalleryAsset
  onClose?: () => void
  onDownload?: (asset: GalleryAsset) => void
  onDelete?: (asset: GalleryAsset) => void
  formatFileSize?: (bytes: number) => string
  formatUploadedAt?: (isoDate: string) => string
}

export function SelectedImagePanel({
  asset,
  onClose,
  onDownload,
  onDelete,
  formatFileSize,
  formatUploadedAt,
}: SelectedImagePanelProps) {
  const { t } = useTranslation()
  const previewUrl = asset.fileUrl ?? asset.thumbnailUrl
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
            alt={asset.altText ?? asset.fileName}
            className={styles.previewImage}
          />
        </div>
      )}

      <div className={styles.heading}>
        <p className={styles.fileName}>{asset.fileName}</p>
        {uploadedLine && <p className={styles.uploadedLine}>{uploadedLine}</p>}
      </div>

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

      {asset.altText && (
        <section className={styles.section}>
          <p className={styles.sectionLabel}>
            {t('galleryManager.selected.altText.title')}
          </p>
          <p className={styles.altStatus}>
            <CheckCircleIcon />
            <span>{t('galleryManager.selected.altText.compliant')}</span>
          </p>
          <p className={styles.altText}>{`"${asset.altText}"`}</p>
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
        <button
          type="button"
          className={styles.downloadButton}
          onClick={() => onDownload?.(asset)}
        >
          <DownloadIcon size={16} />
          {t('galleryManager.selected.actions.download')}
        </button>
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
