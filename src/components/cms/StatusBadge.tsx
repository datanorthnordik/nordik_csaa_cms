import { useTranslation } from 'react-i18next'
import styles from './StatusBadge.module.css'

export type PublishStatus = 'published' | 'draft' | 'scheduled'

type StatusBadgeProps = {
  status: PublishStatus
  size?: 'sm' | 'md' | 'lg'
  label?: string
  className?: string
}

const sizeClass: Record<NonNullable<StatusBadgeProps['size']>, string> = {
  sm: styles.sizeSm,
  md: '',
  lg: styles.sizeLg,
}

export function StatusBadge({
  status,
  size = 'md',
  label,
  className,
}: StatusBadgeProps) {
  const { t } = useTranslation()
  const resolvedLabel = label ?? t(`publishing.status.${status}`)

  const classes = [styles.badge, styles[status], sizeClass[size], className]
    .filter(Boolean)
    .join(' ')

  return <span className={classes}>{resolvedLabel}</span>
}
