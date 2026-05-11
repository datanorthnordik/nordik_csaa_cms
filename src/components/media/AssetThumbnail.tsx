import type { DragEvent } from 'react'
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
  draggable?: boolean
  isDragging?: boolean
  isDragOver?: boolean
  onDragStart?: (asset: GalleryAsset) => void
  onDragEnd?: () => void
  onDragEnter?: (asset: GalleryAsset) => void
  onDragLeave?: (asset: GalleryAsset) => void
  onDragDrop?: (asset: GalleryAsset) => void
}

export function AssetThumbnail({
  asset,
  selected = false,
  onSelect,
  onMoveLeft,
  onMoveRight,
  onDelete,
  draggable = false,
  isDragging = false,
  isDragOver = false,
  onDragStart,
  onDragEnd,
  onDragEnter,
  onDragLeave,
  onDragDrop,
}: AssetThumbnailProps) {
  const { t } = useTranslation()
  const previewUrl = asset.thumbnailUrl ?? asset.fileUrl
  const hasOverlay = Boolean(onMoveLeft || onMoveRight || onDelete)

  function handleDragStart(event: DragEvent<HTMLDivElement>) {
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', String(asset.id))
    onDragStart?.(asset)
  }

  function handleDragOver(event: DragEvent<HTMLDivElement>) {
    if (!draggable) {
      return
    }
    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
  }

  function handleDragEnter(event: DragEvent<HTMLDivElement>) {
    if (!draggable) {
      return
    }
    event.preventDefault()
    onDragEnter?.(asset)
  }

  function handleDragLeave() {
    if (!draggable) {
      return
    }
    onDragLeave?.(asset)
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    if (!draggable) {
      return
    }
    event.preventDefault()
    onDragDrop?.(asset)
  }

  const className = [
    styles.thumbnail,
    selected ? styles.selected : '',
    isDragging ? styles.dragging : '',
    isDragOver ? styles.dragOver : '',
  ]
    .filter(Boolean)
    .join(' ')

  return (
    <div
      className={className}
      draggable={draggable}
      onDragStart={draggable ? handleDragStart : undefined}
      onDragEnd={draggable ? onDragEnd : undefined}
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
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
            draggable={false}
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
