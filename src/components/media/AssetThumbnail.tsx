import { useTranslation } from 'react-i18next'
import {
  ChevronLeftIcon,
  ChevronRightIcon,
  DeleteIcon,
} from '../icons'
import type { GalleryAsset } from '../../types/media'
import styles from './AssetThumbnail.module.css'

type AssetThumbnailProps = {
  asset: GalleryAsset
  selected?: boolean
  onSelect?: (asset: GalleryAsset) => void
  onMoveLeft?: (asset: GalleryAsset) => void
  onMoveRight?: (asset: GalleryAsset) => void
  onDelete?: (asset: GalleryAsset) => void
}

export function AssetThumbnail({
  asset,
  selected = false,
  onSelect,
  onMoveLeft,
  onMoveRight,
  onDelete,
}: AssetThumbnailProps) {
  const { t } = useTranslation()
  const previewUrl = asset.thumbnailUrl ?? asset.fileUrl
  const hasOverlay = Boolean(onMoveLeft || onMoveRight || onDelete)

  return (
    <div
      className={[styles.thumbnail, selected ? styles.selected : '']
        .filter(Boolean)
        .join(' ')}
    >
      <button
        type="button"
        className={styles.selectArea}
        onClick={() => onSelect?.(asset)}
        aria-pressed={selected}
        aria-label={asset.altText ?? asset.fileName}
      >
        {previewUrl && (
          <img
            src={previewUrl}
            alt={asset.altText ?? asset.fileName}
            className={styles.image}
            loading="lazy"
          />
        )}
      </button>

      {selected && (
        <span className={styles.selectedBadge}>
          {t('galleryManager.thumbnail.selected')}
        </span>
      )}

      {hasOverlay && (
        <div className={styles.overlay}>
          {onMoveLeft && (
            <button
              type="button"
              className={styles.overlayButton}
              onClick={() => onMoveLeft(asset)}
              aria-label={t('galleryManager.thumbnail.moveLeft')}
            >
              <ChevronLeftIcon size={16} />
            </button>
          )}
          {onMoveRight && (
            <button
              type="button"
              className={styles.overlayButton}
              onClick={() => onMoveRight(asset)}
              aria-label={t('galleryManager.thumbnail.moveRight')}
            >
              <ChevronRightIcon size={16} />
            </button>
          )}
          {onDelete && (
            <button
              type="button"
              className={[styles.overlayButton, styles.overlayDanger].join(' ')}
              onClick={() => onDelete(asset)}
              aria-label={t('galleryManager.thumbnail.delete')}
            >
              <DeleteIcon size={16} />
            </button>
          )}
        </div>
      )}
    </div>
  )
}
