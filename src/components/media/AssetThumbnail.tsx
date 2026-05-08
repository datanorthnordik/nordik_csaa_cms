import { useTranslation } from 'react-i18next'
import type { GalleryAsset } from '../../types/media'
import styles from './AssetThumbnail.module.css'

type AssetThumbnailProps = {
  asset: GalleryAsset
  selected?: boolean
  onSelect?: (asset: GalleryAsset) => void
}

export function AssetThumbnail({
  asset,
  selected = false,
  onSelect,
}: AssetThumbnailProps) {
  const { t } = useTranslation()
  const previewUrl = asset.thumbnailUrl ?? asset.fileUrl

  return (
    <button
      type="button"
      className={[styles.thumbnail, selected ? styles.selected : '']
        .filter(Boolean)
        .join(' ')}
      onClick={() => onSelect?.(asset)}
      aria-pressed={selected}
    >
      {previewUrl && (
        <img
          src={previewUrl}
          alt={asset.altText ?? asset.fileName}
          className={styles.image}
          loading="lazy"
        />
      )}
      {selected && (
        <span className={styles.selectedBadge}>
          {t('galleryManager.thumbnail.selected')}
        </span>
      )}
    </button>
  )
}
