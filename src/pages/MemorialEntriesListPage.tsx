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
import {
  MOCK_MEMORIAL_ENTRIES,
  defaultMemorialListFilters,
  type MemorialEntry,
  type MemorialListFilters,
  type MemorialStatus,
} from '../lib/memorialTypes'
import styles from '../styles/MemorialEntriesListPage.module.css'

const PAGE_SIZE = 10

export function MemorialEntriesListPage() {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()

  const [entries, setEntries] = useState<MemorialEntry[]>(MOCK_MEMORIAL_ENTRIES)
  const [filters, setFilters] = useState<MemorialListFilters>(defaultMemorialListFilters)
  const [currentPage, setCurrentPage] = useState(1)
  const [deleteCandidate, setDeleteCandidate] = useState<MemorialEntry | null>(null)

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  const filtered = useMemo(() => {
    return entries.filter((e) => {
      if (filters.status !== 'all' && e.status !== filters.status) return false
      if (filters.category && e.category !== filters.category) return false
      if (filters.searchTerm) {
        const q = filters.searchTerm.toLowerCase()
        return (
          e.fullName.toLowerCase().includes(q) ||
          e.affiliation.toLowerCase().includes(q)
        )
      }
      return true
    })
  }, [entries, filters])

  const total = filtered.length
  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE))
  const paginated = filtered.slice(
    (currentPage - 1) * PAGE_SIZE,
    currentPage * PAGE_SIZE,
  )
  const rangeStart = paginated.length ? (currentPage - 1) * PAGE_SIZE + 1 : 0
  const rangeEnd = Math.min(currentPage * PAGE_SIZE, total)

  function updateFilter<K extends keyof MemorialListFilters>(
    key: K,
    value: MemorialListFilters[K],
  ) {
    setFilters((c) => ({ ...c, [key]: value }))
    setCurrentPage(1)
  }

  function formatDate(iso: string) {
    const d = Date.parse(iso)
    return Number.isNaN(d) ? iso : dateFormatter.format(new Date(d))
  }

  function getInitials(name: string) {
    return name
      .split(' ')
      .slice(0, 2)
      .map((w) => w[0] ?? '')
      .join('')
      .toUpperCase()
  }

  function handleConfirmDelete() {
    if (!deleteCandidate) return
    setEntries((c) => c.filter((e) => e.id !== deleteCandidate.id))
    setDeleteCandidate(null)
    toast.success(t('memorial.feedback.deleted'))
  }

  const statusFilterOptions: Array<'all' | MemorialStatus> = [
    'all',
    'published',
    'draft',
  ]

  const fields: FilterFieldConfig[] = [
    {
      type: 'select',
      key: 'status',
      label: t('memorial.filters.status'),
      value: filters.status,
      options: statusFilterOptions.map((v) => ({
        value: v,
        label: t(`memorial.statusFilter.${v}`),
      })),
      onChange: (v) =>
        updateFilter('status', v as MemorialListFilters['status']),
    },
  ]

  return (
    <CmsAppShell activeKey="memorial">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('memorial.breadcrumb.entries') }]} />

        <div className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('memorial.list.title')}</h1>
            <p className={styles.subtitle}>{t('memorial.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/memorial/new')}
          >
            <AddIcon size={16} />
            {t('memorial.list.create')}
          </button>
        </div>

        <SearchFilterBar
          searchValue={filters.searchTerm}
          onSearchChange={(v) => updateFilter('searchTerm', v)}
          searchPlaceholder={t('memorial.filters.searchPlaceholder')}
          searchLabel={t('memorial.filters.search')}
          fields={fields}
          compact
          collapsible
        />

        <section className={styles.resultsPanel}>
          {paginated.length > 0 ? (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>{t('memorial.list.columns.entryName')}</th>
                    <th>{t('memorial.list.columns.status')}</th>
                    <th>{t('memorial.list.columns.dateAdded')}</th>
                    <th>{t('memorial.list.columns.actions')}</th>
                  </tr>
                </thead>
                <tbody>
                  {paginated.map((entry) => (
                    <tr key={entry.id}>
                      <td>
                        <div className={styles.nameCell}>
                          <div className={styles.avatar}>
                            {getInitials(entry.fullName)}
                          </div>
                          <div>
                            <button
                              type="button"
                              className={styles.entryName}
                              onClick={() =>
                                navigate(`/memorial/${entry.id}/edit`)
                              }
                            >
                              {entry.fullName}
                            </button>
                            <span className={styles.entryAffiliation}>
                              {entry.affiliation}
                            </span>
                          </div>
                        </div>
                      </td>
                      <td>
                        <span
                          className={[
                            styles.statusBadge,
                            styles[`status_${entry.status}`] ?? '',
                          ]
                            .filter(Boolean)
                            .join(' ')}
                        >
                          {t(`memorial.status.${entry.status}`)}
                        </span>
                      </td>
                      <td className={styles.dateCell}>
                        {formatDate(entry.dateAdded)}
                      </td>
                      <td className={styles.actionsCell}>
                        <span className={styles.actionsRow}>
                          <button
                            type="button"
                            className={styles.iconButton}
                            aria-label={t('memorial.list.edit')}
                            onClick={() =>
                              navigate(`/memorial/${entry.id}/edit`)
                            }
                          >
                            <PencilIcon />
                          </button>
                          <button
                            type="button"
                            className={`${styles.iconButton} ${styles.iconButtonDanger}`}
                            aria-label={t('memorial.list.delete')}
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
              <p className={styles.emptyTitle}>
                {t('memorial.list.emptyTitle')}
              </p>
              <p>{t('memorial.list.emptyText')}</p>
            </div>
          )}

          {total > 0 && (
            <div className={styles.footer}>
              <span className={styles.footerInfo}>
                {t('memorial.list.resultsLabel', {
                  start: rangeStart,
                  end: rangeEnd,
                  total,
                })}
              </span>
              <PaginationControls
                page={currentPage}
                totalPages={totalPages}
                onChange={setCurrentPage}
              />
            </div>
          )}
        </section>

        <ConfirmDialog
          open={Boolean(deleteCandidate)}
          title={t('memorial.list.deleteDialogTitle')}
          body={
            <Trans
              i18nKey="memorial.list.deleteDialogDescription"
              values={{ title: deleteCandidate?.fullName ?? '' }}
              components={{ entryName: <strong /> }}
            />
          }
          confirmLabel={t('memorial.list.confirmDelete')}
          cancelLabel={t('memorial.list.cancelDelete')}
          destructive
          onConfirm={handleConfirmDelete}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function PencilIcon() {
  return (
    <svg viewBox="0 0 16 16" width="15" height="15" fill="none" aria-hidden="true">
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
    <svg viewBox="0 0 16 16" width="15" height="15" fill="none" aria-hidden="true">
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
