import { useTranslation } from 'react-i18next'
import { DeleteIcon, GalleryIcon } from '../icons'
import type { GallerySummary } from '../../types/media'
import { GalleryStatusChip } from './GalleryStatusChip'
import styles from './GalleryCard.module.css'

type GalleryCardProps = {
  gallery: GallerySummary
  onManage?: (gallery: GallerySummary) => void
  onDelete?: (gallery: GallerySummary) => void
  formatRelativeTime?: (isoDate: string) => string
}

export function GalleryCard({
  gallery,
  onManage,
  onDelete,
  formatRelativeTime,
}: GalleryCardProps) {
  const { t } = useTranslation()

  return (
    <article className={styles.card}>
      <div className={styles.thumbnail}>
        {gallery.frontImageUrl ? (
          <img
            src={gallery.frontImageUrl}
            alt=""
            className={styles.image}
            loading="lazy"
          />
        ) : (
          <div className={styles.placeholder} aria-hidden="true">
            <GalleryIcon size={28} />
          </div>
        )}
      </div>

      <div className={styles.body}>
        <div className={styles.bodyHeader}>
          <GalleryStatusChip
            visibility={gallery.visibility}
            className={styles.statusChip}
          />
          <h3 className={styles.name}>{gallery.name}</h3>
        </div>

        <div className={styles.meta}>
          {typeof gallery.assetCount === 'number' && (
            <span className={styles.metaItem}>
              <GalleryIcon size={14} />
              {t('mediaLibrary.card.assetCount', { count: gallery.assetCount })}
            </span>
          )}
          {gallery.updatedAt && (
            <span className={styles.metaItem}>
              {formatRelativeTime
                ? t('mediaLibrary.card.updatedRelative', {
                    relative: formatRelativeTime(gallery.updatedAt),
                  })
                : t('mediaLibrary.card.updatedAt', {
                    date: gallery.updatedAt,
                  })}
            </span>
          )}
        </div>

        <div className={styles.actions}>
          <button
            type="button"
            className={styles.manageButton}
            onClick={() => onManage?.(gallery)}
          >
            {t('mediaLibrary.card.manage')}
          </button>
          {onDelete && (
            <button
              type="button"
              className={styles.deleteButton}
              onClick={() => onDelete(gallery)}
            >
              <DeleteIcon size={14} />
              {t('mediaLibrary.card.delete')}
            </button>
          )}
        </div>
      </div>
    </article>
  )
}
