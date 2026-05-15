import { useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { AddIcon } from '../components/icons'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { PaginationControls } from '../components/cms/PaginationControls'
import {
  SearchFilterBar,
  type FilterFieldConfig,
} from '../components/cms/SearchFilterBar'
import { StatusBadge } from '../components/cms/StatusBadge'
import { useNewsletterEntries } from '../hooks/useNewsletterEntries'
import {
  defaultNewsletterListFilters,
  type NewsletterEntry,
  type NewsletterListFilters,
  type NewsletterStatus,
} from '../lib/newsletterTypes'
import styles from '../styles/NewslettersListPage.module.css'

const PAGE_SIZE = 10
const statusOptionValues: ('all' | NewsletterStatus)[] = [
  'all',
  'published',
  'draft',
  'scheduled',
]
const sortOptionValues: NewsletterListFilters['sortBy'][] = [
  'sendDate',
  'title',
  'updatedAt',
]

const CATEGORY_LABELS: Record<string, string> = {
  csaa: 'CSAA',
  cst: 'CST',
}

export function NewslettersListPage() {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { entries, remove } = useNewsletterEntries()
  const [filters, setFilters] = useState<NewsletterListFilters>(defaultNewsletterListFilters)
  const [page, setPage] = useState(1)
  const [deleteCandidate, setDeleteCandidate] = useState<NewsletterEntry | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
        timeZone: 'UTC',
      }),
    [i18n.language],
  )

  const filtered = useMemo(() => {
    const term = filters.searchTerm.trim().toLowerCase()
    const filteredList = entries.filter((entry) => {
      if (filters.status !== 'all' && entry.status !== filters.status) {
        return false
      }
      if (term && !entry.title.toLowerCase().includes(term)) {
        return false
      }
      return true
    })

    const sorted = [...filteredList].sort((a, b) => {
      const direction = filters.sortOrder === 'asc' ? 1 : -1
      if (filters.sortBy === 'title') {
        return a.title.localeCompare(b.title) * direction
      }
      if (filters.sortBy === 'updatedAt') {
        return (Date.parse(a.updatedAt) - Date.parse(b.updatedAt)) * direction
      }
      return (Date.parse(a.sendDate) - Date.parse(b.sendDate)) * direction
    })

    return sorted
  }, [entries, filters])

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE))
  const safePage = Math.min(page, totalPages)
  const pageStart = (safePage - 1) * PAGE_SIZE
  const pageItems = filtered.slice(pageStart, pageStart + PAGE_SIZE)
  const rangeStart = filtered.length ? pageStart + 1 : 0
  const rangeEnd = Math.min(pageStart + PAGE_SIZE, filtered.length)

  function updateFilter<K extends keyof NewsletterListFilters>(
    key: K,
    value: NewsletterListFilters[K],
  ) {
    setFilters((current) => ({ ...current, [key]: value }))
    setPage(1)
  }

  function formatDate(iso: string) {
    const parsed = Date.parse(iso)
    if (Number.isNaN(parsed)) {
      return iso
    }
    return dateFormatter.format(new Date(parsed))
  }

  function isModified(entry: NewsletterEntry) {
    return entry.updatedAt !== entry.createdAt
  }

  async function handleConfirmDelete() {
    if (!deleteCandidate) {
      return
    }
    setIsDeleting(true)
    try {
      remove(deleteCandidate.id)
      toast.success(t('newsletters.feedback.deleted'))
      setDeleteCandidate(null)
    } finally {
      setIsDeleting(false)
    }
  }

  const fields: FilterFieldConfig[] = [
    {
      type: 'select',
      key: 'status',
      label: t('newsletters.filters.status'),
      value: filters.status,
      options: statusOptionValues.map((value) => ({
        value,
        label: t(`newsletters.statusFilter.${value}`),
      })),
      onChange: (value) =>
        updateFilter('status', value as NewsletterListFilters['status']),
    },
    {
      type: 'select',
      key: 'sortBy',
      label: t('newsletters.filters.sortBy'),
      value: filters.sortBy,
      options: sortOptionValues.map((value) => ({
        value,
        label: t(`newsletters.sortBy.${value}`),
      })),
      onChange: (value) =>
        updateFilter('sortBy', value as NewsletterListFilters['sortBy']),
    },
  ]

  return (
    <CmsAppShell activeKey="newsletters">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('newsletters.breadcrumb.entries') }]} />

        <div className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('newsletters.list.title')}</h1>
            <p className={styles.subtitle}>{t('newsletters.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/newsletters/new')}
          >
            <AddIcon size={16} />
            {t('newsletters.list.create')}
          </button>
        </div>

        <SearchFilterBar
          searchValue={filters.searchTerm}
          onSearchChange={(value) => updateFilter('searchTerm', value)}
          searchPlaceholder={t('newsletters.filters.searchPlaceholder')}
          searchLabel={t('newsletters.filters.search')}
          fields={fields}
          compact
        />

        <section className={styles.resultsPanel}>
          {filtered.length > 0 && (
            <div className={styles.footer}>
              <span className={styles.footerInfo}>
                {t('newsletters.list.resultsLabel', {
                  start: rangeStart,
                  end: rangeEnd,
                  total: filtered.length,
                })}
              </span>
              <PaginationControls
                page={safePage}
                totalPages={totalPages}
                onChange={setPage}
              />
            </div>
          )}

          {pageItems.length ? (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>{t('newsletters.list.columns.title')}</th>
                    <th>{t('newsletters.list.columns.sendDate')}</th>
                    <th>{t('newsletters.list.columns.status')}</th>
                    <th>{t('newsletters.list.columns.actions')}</th>
                  </tr>
                </thead>
                <tbody>
                  {pageItems.map((entry) => (
                    <tr
                      key={entry.id}
                    >
                      <td>
                        <div className={styles.titleCell}>
                          {entry.category && (
                            <span
                              className={`${styles.categoryChip} ${styles[`categoryChip_${entry.category}`]}`}
                              aria-hidden="true"
                            >
                              {CATEGORY_LABELS[entry.category]}
                            </span>
                          )}
                          <div>
                            <button
                              type="button"
                              className={styles.entryTitle}
                              onClick={() => navigate(`/newsletters/${entry.id}/edit`)}
                            >
                              {entry.title || t('newsletters.list.untitled')}
                            </button>
                            {isModified(entry) && (
                              <span className={styles.modifiedHint}>
                                {t('newsletters.list.modifiedOn', {
                                  date: formatDate(entry.updatedAt),
                                })}
                              </span>
                            )}
                          </div>
                        </div>
                      </td>
                      <td className={styles.dateCell}>
                        {formatDate(entry.sendDate)}
                      </td>
                      <td>
                        <StatusBadge status={entry.status} />
                      </td>
                      <td className={styles.actionsCell}>
                        <span className={styles.actionsRow}>
                          <button
                            type="button"
                            className={styles.iconButton}
                            aria-label={t('newsletters.list.edit')}
                            onClick={() => navigate(`/newsletters/${entry.id}/edit`)}
                          >
                            <PencilIcon />
                          </button>
                          <button
                            type="button"
                            className={`${styles.iconButton} ${styles.iconButtonDanger}`}
                            aria-label={t('newsletters.list.delete')}
                            onClick={() => setDeleteCandidate(entry)}
                          >
                            <TrashIcon />
                          </button>
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>{t('newsletters.list.emptyTitle')}</p>
              <p>{t('newsletters.list.emptyText')}</p>
            </div>
          )}

        </section>

        <ConfirmDialog
          open={Boolean(deleteCandidate)}
          title={t('newsletters.list.deleteDialogTitle')}
          body={
            <Trans
              i18nKey="newsletters.list.deleteDialogDescription"
              values={{ title: deleteCandidate?.title ?? '' }}
              components={{ entryName: <strong /> }}
            />
          }
          confirmLabel={t('newsletters.list.delete')}
          cancelLabel={t('newsletters.list.cancelDelete')}
          destructive
          busy={isDeleting}
          onConfirm={() => void handleConfirmDelete()}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function PencilIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M10.6 2.6 13.4 5.4M2.6 13.4l8.4-8.4 2.6 2.6-8.4 8.4-3 .4z"
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
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
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
