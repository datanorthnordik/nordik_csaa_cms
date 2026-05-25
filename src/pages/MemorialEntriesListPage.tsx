import { useDeferredValue, useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { AddIcon } from '../components/icons'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { PaginationControls } from '../components/cms/PaginationControls'
import {
  SearchFilterBar,
  type FilterFieldConfig,
} from '../components/cms/SearchFilterBar'
import { useMemorialEntries } from '../hooks/useMemorialEntries'
import {
  MEMORIAL_CATEGORIES,
  defaultMemorialListFilters,
  type MemorialCategory,
  type MemorialEntrySummary,
  type MemorialListFilters,
  type MemorialStatus,
} from '../lib/memorialTypes'
import styles from '../styles/MemorialEntriesListPage.module.css'

const PAGE_SIZE = 10

export function MemorialEntriesListPage() {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { items, pagination, loading, error, fetch, remove } = useMemorialEntries(PAGE_SIZE)
  const [filters, setFilters] = useState<MemorialListFilters>({
    ...defaultMemorialListFilters,
    pageSize: PAGE_SIZE,
  })
  const [deleteCandidate, setDeleteCandidate] = useState<MemorialEntrySummary | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  const deferredSearchTerm = useDeferredValue(filters.searchTerm)

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  useEffect(() => {
    void fetch({
      ...filters,
      searchTerm: deferredSearchTerm,
    }).catch((fetchError) => {
      toast.error(
        fetchError instanceof Error
          ? fetchError.message
          : t('memorial.feedback.fetchError'),
      )
    })
  }, [
    deferredSearchTerm,
    fetch,
    filters.category,
    filters.page,
    filters.pageSize,
    filters.status,
    t,
  ])

  function updateFilters(patch: Partial<MemorialListFilters>) {
    setFilters((current) => ({
      ...current,
      ...patch,
    }))
  }

  function formatDate(value: string) {
    const parsed = Date.parse(value)
    if (Number.isNaN(parsed)) {
      return value
    }
    return dateFormatter.format(new Date(parsed))
  }

  function getInitials(name: string) {
    return name
      .split(' ')
      .slice(0, 2)
      .map((word) => word[0] ?? '')
      .join('')
      .toUpperCase()
  }

  async function handleConfirmDelete() {
    if (!deleteCandidate) {
      return
    }

    setIsDeleting(true)
    try {
      await remove(deleteCandidate.id)
      const nextPage =
        pagination.page > 1 && items.length === 1
          ? pagination.page - 1
          : pagination.page

      await fetch({
        ...filters,
        page: nextPage,
        searchTerm: deferredSearchTerm,
      })

      setFilters((current) => ({
        ...current,
        page: nextPage,
      }))
      setDeleteCandidate(null)
      toast.success(t('memorial.feedback.deleted'))
    } catch (deleteError) {
      toast.error(
        deleteError instanceof Error
          ? deleteError.message
          : t('memorial.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  const totalItems = pagination.totalItems
  const rangeStart =
    totalItems > 0 ? (pagination.page - 1) * pagination.pageSize + 1 : 0
  const rangeEnd =
    totalItems > 0
      ? Math.min(pagination.page * pagination.pageSize, pagination.totalItems)
      : 0

  const statusFilterOptions: Array<'all' | MemorialStatus> = [
    'all',
    'published',
    'draft',
    'review',
  ]

  const fields: FilterFieldConfig[] = [
    {
      type: 'select',
      key: 'status',
      label: t('memorial.filters.status'),
      value: filters.status,
      options: statusFilterOptions.map((value) => ({
        value,
        label: t(`memorial.statusFilter.${value}`),
      })),
      onChange: (value) =>
        updateFilters({
          status: value as MemorialListFilters['status'],
          page: 1,
        }),
    },
    {
      type: 'select',
      key: 'category',
      label: t('memorial.filters.category'),
      value: filters.category,
      options: [
        {
          value: '',
          label: t('memorial.filters.allCategories'),
        },
        ...MEMORIAL_CATEGORIES.map((category) => ({
          value: category,
          label: t(`memorial.category.${category}`),
        })),
      ],
      onChange: (value) =>
        updateFilters({
          category: value as MemorialCategory | '',
          page: 1,
        }),
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
          onSearchChange={(value) =>
            updateFilters({ searchTerm: value, page: 1 })
          }
          searchPlaceholder={t('memorial.filters.searchPlaceholder')}
          searchLabel={t('memorial.filters.search')}
          fields={fields}
          compact
          collapsible
        />

        <section className={styles.resultsPanel}>
          {error && (
            <div className={styles.errorBox}>
              <p>{error}</p>
            </div>
          )}

          {loading ? (
            <div className={styles.loadingContainer}>
              <Loader label={t('memorial.common.loading')} />
            </div>
          ) : items.length > 0 ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>{t('memorial.list.columns.entryName')}</th>
                      <th>{t('memorial.list.columns.category')}</th>
                      <th>{t('memorial.list.columns.status')}</th>
                      <th>{t('memorial.list.columns.dateAdded')}</th>
                      <th>{t('memorial.list.columns.actions')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((entry) => (
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
                              styles.categoryBadge,
                              styles[`category_${entry.category}`] ?? '',
                            ]
                              .filter(Boolean)
                              .join(' ')}
                          >
                            {entry.categoryLabel}
                          </span>
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
                          {formatDate(entry.createdAt)}
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

              <div className={styles.footer}>
                <span className={styles.footerInfo}>
                  {t('memorial.list.resultsLabel', {
                    start: rangeStart,
                    end: rangeEnd,
                    total: totalItems,
                  })}
                </span>
                <PaginationControls
                  page={pagination.page}
                  totalPages={pagination.totalPages}
                  onChange={(page) => updateFilters({ page })}
                />
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>
                {t('memorial.list.emptyTitle')}
              </p>
              <p>{t('memorial.list.emptyText')}</p>
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
