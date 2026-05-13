import { useTranslation } from 'react-i18next'
import type { MediaVisibility } from '../../types/media'
import styles from './GalleryStatusChip.module.css'

type GalleryStatusChipProps = {
  visibility?: MediaVisibility
  className?: string
}

export function GalleryStatusChip({
  visibility = 'draft',
  className,
}: GalleryStatusChipProps) {
  const { t } = useTranslation()

  return (
    <span
      className={[styles.chip, styles[visibility], className]
        .filter(Boolean)
        .join(' ')}
    >
      {t(`mediaLibrary.visibility.${visibility}`)}
    </span>
  )
}
