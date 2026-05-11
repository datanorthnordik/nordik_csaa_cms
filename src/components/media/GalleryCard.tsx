import { useTranslation } from 'react-i18next'
import { GalleryIcon } from '../icons'
import type { GallerySummary } from '../../types/media'
import styles from './GalleryCard.module.css'

type GalleryCardProps = {
  gallery: GallerySummary
  onManage?: (gallery: GallerySummary) => void
  formatRelativeTime?: (isoDate: string) => string
}

export function GalleryCard({
  gallery,
  onManage,
  formatRelativeTime,
}: GalleryCardProps) {
  const { t } = useTranslation()
  const visibilityLabel = gallery.visibility
    ? t(`mediaLibrary.visibility.${gallery.visibility}`)
    : null

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
        {visibilityLabel && (
          <span className={styles.visibility}>{visibilityLabel}</span>
        )}
      </div>

      <div className={styles.body}>
        <h3 className={styles.name}>{gallery.name}</h3>

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

        <button
          type="button"
          className={styles.manageButton}
          onClick={() => onManage?.(gallery)}
        >
          {t('mediaLibrary.card.manage')}
        </button>
      </div>
    </article>
  )
}
