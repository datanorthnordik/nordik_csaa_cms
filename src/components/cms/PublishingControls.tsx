import type { ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { StatusBadge, type PublishStatus } from './StatusBadge'
import styles from './PublishingControls.module.css'

export type PublishVisibility = 'public' | 'private'

export type PublishingControlsProps = {
  status: PublishStatus
  visibility?: PublishVisibility
  onVisibilityChange?: (value: PublishVisibility) => void
  publishOn?: string
  heading?: string
  extraSlot?: ReactNode
  className?: string
}

export function PublishingControls({
  status,
  visibility,
  onVisibilityChange,
  publishOn,
  heading,
  extraSlot,
  className,
}: PublishingControlsProps) {
  const { t } = useTranslation()

  const resolvedHeading = heading ?? t('publishing.heading')
  const resolvedPublishOn = publishOn ?? t('publishing.immediately')

  const cardClasses = [styles.card, className].filter(Boolean).join(' ')

  return (
    <aside className={cardClasses} aria-labelledby="publishing-heading">
      <div className={styles.header}>
        <h2 id="publishing-heading" className={styles.heading}>
          {resolvedHeading}
        </h2>
      </div>

      <div className={styles.metaList}>
        <div className={styles.metaRow}>
          <span className={styles.metaLabel}>
            <span className={styles.metaIcon} aria-hidden="true">
              <StatusDotIcon />
            </span>
            {t('publishing.statusLabel')}
          </span>
          <StatusBadge status={status} />
        </div>

        {visibility !== undefined && (
          <div className={styles.metaRow}>
            <span className={styles.metaLabel}>
              <span className={styles.metaIcon} aria-hidden="true">
                <EyeIcon />
              </span>
              {t('publishing.visibilityLabel')}
            </span>
            {onVisibilityChange ? (
              <select
                className={styles.visibilitySelect}
                value={visibility}
                onChange={(event) =>
                  onVisibilityChange(event.target.value as PublishVisibility)
                }
                aria-label={t('publishing.visibilityLabel')}
              >
                <option value="public">{t('publishing.visibility.public')}</option>
                <option value="private">{t('publishing.visibility.private')}</option>
              </select>
            ) : (
              <span className={styles.metaValue}>
                {t(`publishing.visibility.${visibility}`)}
              </span>
            )}
          </div>
        )}

        <div className={styles.metaRow}>
          <span className={styles.metaLabel}>
            <span className={styles.metaIcon} aria-hidden="true">
              <ClockIcon />
            </span>
            {t('publishing.publishOnLabel')}
          </span>
          <span className={styles.metaValue}>{resolvedPublishOn}</span>
        </div>
      </div>

      {extraSlot && <div className={styles.extras}>{extraSlot}</div>}
    </aside>
  )
}

function StatusDotIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6" stroke="currentColor" strokeWidth="1.6" />
      <circle cx="8" cy="8" r="2.4" fill="currentColor" />
    </svg>
  )
}

function EyeIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M1.6 8c1.6-3.1 4-4.7 6.4-4.7S12.8 4.9 14.4 8c-1.6 3.1-4 4.7-6.4 4.7S3.2 11.1 1.6 8Z"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinejoin="round"
      />
      <circle cx="8" cy="8" r="2" stroke="currentColor" strokeWidth="1.4" />
    </svg>
  )
}

function ClockIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6.2" stroke="currentColor" strokeWidth="1.4" />
      <path
        d="M8 4.6V8l2.4 1.4"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function TrashIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M2.8 4.4h10.4M6 4.4V3.2c0-.66.54-1.2 1.2-1.2h1.6c.66 0 1.2.54 1.2 1.2v1.2M4.4 4.4l.6 8.2c.04.66.6 1.2 1.26 1.2h3.48c.66 0 1.22-.54 1.26-1.2l.6-8.2"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}
