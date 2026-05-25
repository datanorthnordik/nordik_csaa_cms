import { useDeferredValue, useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { resourcesApi } from '../api/resourcesApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, DownloadIcon } from '../components/icons'
import { useResources } from '../hooks/useResources'
import {
  defaultResourceListFilters,
  resourceCategoryOptions,
  resourceFileTypeOptions,
  type ResourceCategory,
  type ResourceEntry,
  type ResourceListFilters,
} from '../lib/resourceTypes'
import styles from '../styles/ResourcesListPage.module.css'

const PAGE_SIZE = 6

export function ResourcesListPage() {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { items, pagination, categoryCounts, loading, error, fetch, remove } = useResources(PAGE_SIZE)
  const [filters, setFilters] = useState<ResourceListFilters>({
    ...defaultResourceListFilters,
    pageSize: PAGE_SIZE,
  })
  const [deleteCandidate, setDeleteCandidate] = useState<ResourceEntry | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  const [activeDownloadId, setActiveDownloadId] = useState<string | null>(null)
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
          : t('resources.feedback.fetchError'),
      )
    })
  }, [deferredSearchTerm, fetch, filters.category, filters.fileType, filters.page, filters.pageSize, t])

  const categoryCountMap = useMemo(() => {
    const counts = new Map<ResourceCategory, number>()
    for (const category of categoryCounts) {
      counts.set(category.category, category.count)
    }
    return counts
  }, [categoryCounts])

  function updateFilters(patch: Partial<ResourceListFilters>) {
    setFilters((current) => ({
      ...current,
      ...patch,
    }))
  }

  function handleCategorySelect(category: ResourceCategory) {
    updateFilters({
      category: filters.category === category ? '' : category,
      page: 1,
    })
  }

  async function handleDownload(item: ResourceEntry) {
    setActiveDownloadId(item.id)
    try {
      const blob = await resourcesApi.getResourceContent(item.id)
      const objectUrl = URL.createObjectURL(blob)
      triggerFileDownload(objectUrl, item.fileName)
      scheduleObjectUrlRevoke(objectUrl)
    } catch {
      toast.error(t('resources.feedback.downloadError'))
    } finally {
      setActiveDownloadId(null)
    }
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
      toast.success(t('resources.feedback.deleted'))
    } catch (deleteError) {
      toast.error(
        deleteError instanceof Error
          ? deleteError.message
          : t('resources.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  function formatDate(value: string) {
    const parsed = Date.parse(value)
    if (Number.isNaN(parsed)) {
      return value
    }
    return dateFormatter.format(new Date(parsed))
  }

  function formatFileSize(size: number) {
    if (size < 1024) {
      return `${size} B`
    }
    if (size < 1024 * 1024) {
      return `${(size / 1024).toFixed(1)} KB`
    }
    return `${(size / (1024 * 1024)).toFixed(1)} MB`
  }

  const totalItems = pagination.totalItems
  const rangeStart = totalItems > 0 ? (pagination.page - 1) * pagination.pageSize + 1 : 0
  const rangeEnd =
    totalItems > 0
      ? Math.min(pagination.page * pagination.pageSize, pagination.totalItems)
      : 0

  return (
    <CmsAppShell activeKey="resources">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('resources.breadcrumb.library') }]} />

        <div className={styles.pageHeader}>
          <h1 className={styles.title}>{t('resources.list.title')}</h1>
          <button
            type="button"
            className={styles.uploadButton}
            onClick={() => navigate('/resources/new')}
          >
            <AddIcon size={16} />
            {t('resources.list.create')}
          </button>
        </div>

        <div className={styles.searchRow}>
          <label className={styles.searchField}>
            <span className={styles.searchIcon} aria-hidden="true">
              <SearchIcon />
            </span>
            <input
              type="search"
              value={filters.searchTerm}
              placeholder={t('resources.filters.searchPlaceholder')}
              aria-label={t('resources.filters.searchLabel')}
              onChange={(event) =>
                updateFilters({ searchTerm: event.target.value, page: 1 })
              }
            />
          </label>

          <label className={styles.typeFilter}>
            <span className={styles.filterIcon} aria-hidden="true">
              <FilterIcon />
            </span>
            <select
              value={filters.fileType}
              aria-label={t('resources.filters.typeLabel')}
              onChange={(event) =>
                updateFilters({
                  fileType: event.target.value as ResourceListFilters['fileType'],
                  page: 1,
                })
              }
            >
              {resourceFileTypeOptions.map((option) => (
                <option key={option.value} value={option.value}>
                  {t(`resources.filters.fileTypes.${option.value}`)}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className={styles.categoryGrid}>
          {resourceCategoryOptions.map((category) => {
            const selected = filters.category === category.id
            const count = categoryCountMap.get(category.id) ?? 0

            return (
              <button
                key={category.id}
                type="button"
                className={`${styles.categoryCard} ${selected ? styles.categoryCardActive : ''}`}
                onClick={() => handleCategorySelect(category.id)}
              >
                <span className={styles.categoryIcon}>
                  <CategoryGlyph category={category.id} />
                </span>
                <span className={styles.categoryTitle}>{category.label}</span>
                <span className={styles.categoryMeta}>
                  {t('resources.list.assetCount', { count })}
                </span>
              </button>
            )
          })}
        </div>

        <section className={styles.resultsPanel}>
          <div className={styles.resultsHeader}>
            <h2>{t('resources.list.recentFiles')}</h2>
            <button
              type="button"
              className={styles.viewAllButton}
              onClick={() => updateFilters({ category: '', page: 1 })}
            >
              {t('resources.list.viewAll')}
            </button>
          </div>

          {error && (
            <div className={styles.errorBox}>
              <p>{error}</p>
            </div>
          )}

          {loading ? (
            <div className={styles.loadingWrap}>
              <Loader label={t('resources.common.loading')} />
            </div>
          ) : items.length > 0 ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>{t('resources.list.columns.name')}</th>
                      <th>{t('resources.list.columns.category')}</th>
                      <th>{t('resources.list.columns.modified')}</th>
                      <th>{t('resources.list.columns.size')}</th>
                      <th>{t('resources.list.columns.actions')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr key={item.id}>
                        <td>
                          <div className={styles.nameCell}>
                            <span className={styles.fileBadge}>
                              {resolveFileBadge(item.fileName)}
                            </span>
                            <div className={styles.nameStack}>
                              <span className={styles.fileTitle}>{item.name}</span>
                              <span className={styles.fileSubline}>{item.fileName}</span>
                            </div>
                          </div>
                        </td>
                        <td>
                          <span className={styles.categoryPill}>{item.categoryLabel}</span>
                        </td>
                        <td>{formatDate(item.updatedAt)}</td>
                        <td>{formatFileSize(item.fileSize)}</td>
                        <td>
                          <div className={styles.actionsCell}>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={t('resources.list.download')}
                              disabled={activeDownloadId === item.id}
                              onClick={() => void handleDownload(item)}
                            >
                              <DownloadIcon size={16} />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={t('resources.list.edit')}
                              onClick={() => navigate(`/resources/${item.id}/edit`)}
                            >
                              <PencilIcon />
                            </button>
                            <button
                              type="button"
                              className={`${styles.iconButton} ${styles.iconButtonDanger}`}
                              aria-label={t('resources.list.delete')}
                              onClick={() => setDeleteCandidate(item)}
                            >
                              <TrashIcon />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className={styles.footer}>
                <span className={styles.footerInfo}>
                  {t('resources.list.resultsLabel', {
                    start: rangeStart,
                    end: rangeEnd,
                    total: totalItems,
                  })}
                </span>

                <div className={styles.pager}>
                  <button
                    type="button"
                    className={styles.pagerButton}
                    disabled={!pagination.hasPrev}
                    onClick={() => updateFilters({ page: pagination.page - 1 })}
                    aria-label={t('pagination.previous')}
                  >
                    <ChevronLeftIcon />
                  </button>
                  <button
                    type="button"
                    className={styles.pagerButton}
                    disabled={!pagination.hasNext}
                    onClick={() => updateFilters({ page: pagination.page + 1 })}
                    aria-label={t('pagination.next')}
                  >
                    <ChevronRightIcon />
                  </button>
                </div>
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>{t('resources.list.emptyTitle')}</p>
              <p>{t('resources.list.emptyText')}</p>
            </div>
          )}
        </section>

        <ConfirmDialog
          open={Boolean(deleteCandidate)}
          title={t('resources.list.deleteDialogTitle')}
          body={t('resources.list.deleteDialogDescription', {
            title: deleteCandidate?.name ?? '',
          })}
          confirmLabel={t('resources.list.delete')}
          cancelLabel={t('resources.list.cancelDelete')}
          destructive
          busy={isDeleting}
          onConfirm={() => void handleConfirmDelete()}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function resolveFileBadge(fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (!extension) {
    return 'FILE'
  }
  return extension.length > 5 ? extension.slice(0, 5) : extension
}

function triggerFileDownload(url: string, fileName: string) {
  const link = globalThis.document.createElement('a')
  link.href = url
  link.download = fileName
  link.rel = 'noreferrer'
  globalThis.document.body.appendChild(link)
  link.click()
  link.remove()
}

function scheduleObjectUrlRevoke(url: string) {
  globalThis.setTimeout(() => {
    URL.revokeObjectURL(url)
  }, 60_000)
}

function SearchIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <circle cx="7" cy="7" r="4.8" stroke="currentColor" strokeWidth="1.4" />
      <path d="M10.8 10.8 14 14" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  )
}

function FilterIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M2.5 4h11M4.5 8h7M6.5 12h3"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
      />
    </svg>
  )
}

function CategoryGlyph({ category }: { category: ResourceCategory }) {
  if (category === 'brand_identity') {
    return (
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
        <circle cx="8" cy="8" r="4.2" stroke="currentColor" strokeWidth="1.5" />
        <circle cx="16" cy="16" r="4.2" stroke="currentColor" strokeWidth="1.5" />
        <path d="M11 11 13 13" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
      </svg>
    )
  }
  if (category === 'governance_legal') {
    return (
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
        <path d="M12 4v14M6 8h12M8 8l-3 6h6L8 8Zm8 0-3 6h6l-3-6Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    )
  }
  if (category === 'training_manuals') {
    return (
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
        <path d="M5 6.5 12 3l7 3.5v9L12 19l-7-3.5v-9Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
        <path d="M12 8.2v6.6M8.7 10 12 11.8 15.3 10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    )
  }
  return (
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
      <path d="M5 5h14v14H5z" stroke="currentColor" strokeWidth="1.5" />
      <path d="M8 15.5V8.5h8v7" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
      <path d="M8 13h8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
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

function ChevronLeftIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path d="M9.8 3.5 5.3 8l4.5 4.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

function ChevronRightIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path d="m6.2 3.5 4.5 4.5-4.5 4.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}
