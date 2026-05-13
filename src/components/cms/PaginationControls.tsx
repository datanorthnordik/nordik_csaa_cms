import { useTranslation } from 'react-i18next'
import styles from './PaginationControls.module.css'

type PaginationControlsProps = {
  page: number
  totalPages: number
  onChange: (page: number) => void
  previousLabel?: string
  nextLabel?: string
  className?: string
}

export function PaginationControls({
  page,
  totalPages,
  onChange,
  previousLabel,
  nextLabel,
  className,
}: PaginationControlsProps) {
  const { t } = useTranslation()
  const resolvedPrev = previousLabel ?? t('pagination.previous')
  const resolvedNext = nextLabel ?? t('pagination.next')

  if (totalPages <= 1) {
    return null
  }

  const pages = buildPageList(page, totalPages)
  const classes = [styles.pagination, className].filter(Boolean).join(' ')

  return (
    <span className={classes}>
      <button
        type="button"
        className={styles.pageButton}
        disabled={page <= 1}
        onClick={() => onChange(page - 1)}
        aria-label={resolvedPrev}
      >
        ‹
      </button>
      {pages.map((entry, index) =>
        entry === 'ellipsis' ? (
          <span key={`ellipsis-${index}`} className={styles.pageEllipsis}>
            …
          </span>
        ) : (
          <button
            key={entry}
            type="button"
            className={`${styles.pageButton} ${
              entry === page ? styles.pageButtonActive : ''
            }`}
            aria-current={entry === page ? 'page' : undefined}
            onClick={() => onChange(entry)}
          >
            {entry}
          </button>
        ),
      )}
      <button
        type="button"
        className={styles.pageButton}
        disabled={page >= totalPages}
        onClick={() => onChange(page + 1)}
        aria-label={resolvedNext}
      >
        ›
      </button>
    </span>
  )
}

function buildPageList(current: number, total: number): (number | 'ellipsis')[] {
  if (total <= 7) {
    return Array.from({ length: total }, (_, i) => i + 1)
  }
  const set = new Set<number>([1, total, current, current - 1, current + 1])
  const sorted = Array.from(set)
    .filter((n) => n >= 1 && n <= total)
    .sort((a, b) => a - b)
  const result: (number | 'ellipsis')[] = []
  sorted.forEach((value, index) => {
    if (index > 0 && value - sorted[index - 1] > 1) {
      result.push('ellipsis')
    }
    result.push(value)
  })
  return result
}
