import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { fetchPressCoverImageContent } from '../api/pressApi'
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
import { StatusBadge } from '../components/cms/StatusBadge'
import { usePressEntries } from '../hooks/usePressEntries'
import {
  defaultPressListFilters,
  type PressEntry,
  type PressListFilters,
  type PressStatus,
} from '../lib/pressTypes'
import { API_BASE_URL } from '../constants/api'
import styles from '../styles/PressListPage.module.css'

const PAGE_SIZE = 10
const statusOptionValues: ('all' | PressStatus)[] = [
  'all',
  'published',
  'draft',
  'scheduled',
]

export function PressListPage() {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const {
    entries,
    loading,
    error,
    totalPages,
    total,
    fetch,
    remove,
  } = usePressEntries(1, PAGE_SIZE)
  
  const [filters, setFilters] = useState<PressListFilters>(defaultPressListFilters)
  const [currentPage, setCurrentPage] = useState(1)
  const [deleteCandidate, setDeleteCandidate] = useState<PressEntry | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  const [coverPreviewUrls, setCoverPreviewUrls] = useState<Record<string, string>>({})
  const coverPreviewUrlsRef = useRef<Record<string, string>>({})

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

  useEffect(() => {
    const entriesNeedingFetch = entries.filter(
      (entry) => entry.coverImageUrl && !resolveDirectPressAssetUrl(entry.coverImageUrl),
    )

    if (!entriesNeedingFetch.length) {
      setCoverPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })
        coverPreviewUrlsRef.current = {}
        return {}
      })
      return
    }

    let cancelled = false

    async function loadCoverPreviews() {
      const nextEntries = await Promise.all(
        entriesNeedingFetch.map(async (entry) => {
          try {
            const blob = await fetchPressCoverImageContent(entry.id)
            return [entry.id, URL.createObjectURL(blob)] as const
          } catch {
            return [entry.id, ''] as const
          }
        }),
      )

      if (cancelled) {
        nextEntries.forEach(([, url]) => {
          if (url) {
            URL.revokeObjectURL(url)
          }
        })
        return
      }

      setCoverPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })

        const next = Object.fromEntries(nextEntries.filter(([, url]) => Boolean(url)))
        coverPreviewUrlsRef.current = next
        return next
      })
    }

    void loadCoverPreviews()

    return () => {
      cancelled = true
    }
  }, [entries])

  useEffect(() => {
    return () => {
      Object.values(coverPreviewUrlsRef.current).forEach((url) => {
        URL.revokeObjectURL(url)
      })
      coverPreviewUrlsRef.current = {}
    }
  }, [])

  // Fetch data when filters or page changes
  useEffect(() => {
    const apiFilters = {
      status: filters.status !== 'all' ? filters.status : undefined,
      search: filters.searchTerm || undefined,
    }
    void fetch(currentPage, PAGE_SIZE, apiFilters).catch((err) => {
      toast.error(
        err instanceof Error
          ? err.message
          : t('press.feedback.fetchError'),
      )
    })
  }, [filters, currentPage, fetch, t])

  function updateFilter<K extends keyof PressListFilters>(
    key: K,
    value: PressListFilters[K],
  ) {
    setFilters((current) => ({ ...current, [key]: value }))
    setCurrentPage(1)
  }

  function formatDate(iso: string) {
    const parsed = Date.parse(iso)
    if (Number.isNaN(parsed)) {
      return iso
    }
    return dateFormatter.format(new Date(parsed))
  }

  function isModified(entry: PressEntry) {
    return entry.updatedAt !== entry.createdAt
  }

  async function handleConfirmDelete() {
    if (!deleteCandidate) {
      return
    }
    setIsDeleting(true)
    try {
      await remove(deleteCandidate.id)
      toast.success(t('press.feedback.deleted'))
      setDeleteCandidate(null)
    } catch (err) {
      toast.error(
        err instanceof Error ? err.message : t('press.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  const rangeStart = entries.length ? (currentPage - 1) * PAGE_SIZE + 1 : 0
  const rangeEnd = Math.min(currentPage * PAGE_SIZE, total)

  const fields: FilterFieldConfig[] = [
    {
      type: 'select',
      key: 'status',
      label: t('press.filters.status'),
      value: filters.status,
      options: statusOptionValues.map((value) => ({
        value,
        label: t(`press.statusFilter.${value}`),
      })),
      onChange: (value) =>
        updateFilter('status', value as PressListFilters['status']),
    },
  ]

  return (
    <CmsAppShell activeKey="press">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('press.breadcrumb.entries') }]} />

        <div className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('press.list.title')}</h1>
            <p className={styles.subtitle}>{t('press.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/press/new')}
          >
            <AddIcon size={16} />
            {t('press.list.create')}
          </button>
        </div>

        <SearchFilterBar
          searchValue={filters.searchTerm}
          onSearchChange={(value) => updateFilter('searchTerm', value)}
          searchPlaceholder={t('press.filters.searchPlaceholder')}
          searchLabel={t('press.filters.search')}
          fields={fields}
          compact
        />

        <section className={styles.resultsPanel}>
          {error && (
            <div className={styles.errorBox}>
              <p>{error}</p>
            </div>
          )}

          {loading ? (
            <div className={styles.loadingContainer}>
              <Loader />
            </div>
          ) : entries.length > 0 ? (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>{t('press.list.columns.title')}</th>
                    <th>{t('press.list.columns.publishDate')}</th>
                    <th>{t('press.list.columns.status')}</th>
                    <th>{t('press.list.columns.actions')}</th>
                  </tr>
                </thead>
                <tbody>
                  {entries.map((entry) => {
                    const coverSrc =
                      resolveDirectPressAssetUrl(entry.coverImageUrl) ||
                      coverPreviewUrls[entry.id] ||
                      ''

                    return (
                      <tr
                        key={entry.id}
                        className={
                          entry.status === 'draft' ? styles.draftRow : undefined
                        }
                      >
                      <td>
                        <div className={styles.titleCell}>
                          <span
                            className={styles.thumbnail}
                            aria-hidden={!coverSrc}
                          >
                            {coverSrc && (
                              <img
                                src={coverSrc}
                                alt=""
                              />
                            )}
                          </span>
                          <div>
                            <button
                              type="button"
                              className={styles.entryTitle}
                              onClick={() => navigate(`/press/${entry.id}/edit`)}
                            >
                              {entry.title || t('press.list.untitled')}
                            </button>
                            {isModified(entry) && (
                              <span className={styles.modifiedHint}>
                                {t('press.list.modifiedOn', {
                                  date: formatDate(entry.updatedAt),
                                })}
                              </span>
                            )}
                          </div>
                        </div>
                      </td>
                      <td className={styles.dateCell}>
                        {formatDate(entry.releaseDate)}
                      </td>
                      <td>
                        <StatusBadge status={entry.status} />
                      </td>
                      <td className={styles.actionsCell}>
                        <span className={styles.actionsRow}>
                          <button
                            type="button"
                            className={styles.iconButton}
                            aria-label={t('press.list.edit')}
                            onClick={() => navigate(`/press/${entry.id}/edit`)}
                          >
                            <PencilIcon />
                          </button>
                          <button
                            type="button"
                            className={`${styles.iconButton} ${styles.iconButtonDanger}`}
                            aria-label={t('press.list.delete')}
                            onClick={() => setDeleteCandidate(entry)}
                          >
                            <TrashIcon />
                          </button>
                        </span>
                      </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>{t('press.list.emptyTitle')}</p>
              <p>{t('press.list.emptyText')}</p>
            </div>
          )}

          {total > 0 && (
            <div className={styles.footer}>
              <span className={styles.footerInfo}>
                {t('press.list.resultsLabel', {
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
          title={t('press.list.deleteDialogTitle')}
          body={
            <Trans
              i18nKey="press.list.deleteDialogDescription"
              values={{ title: deleteCandidate?.title ?? '' }}
              components={{ entryName: <strong /> }}
            />
          }
          confirmLabel={t('press.list.delete')}
          cancelLabel={t('press.list.cancelDelete')}
          destructive
          busy={isDeleting}
          onConfirm={() => void handleConfirmDelete()}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function resolveDirectPressAssetUrl(value?: string) {
  const trimmed = value?.trim() ?? ''
  if (!trimmed) {
    return ''
  }

  if (trimmed.startsWith('blob:') || trimmed.startsWith('data:')) {
    return trimmed
  }

  if (looksLikeManagedStorageReference(trimmed)) {
    return ''
  }

  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed) || trimmed.startsWith('//')) {
    return trimmed
  }

  try {
    return new URL(trimmed, `${API_BASE_URL}/`).toString()
  } catch {
    return trimmed
  }
}

function looksLikeManagedStorageReference(value: string) {
  const normalized = value.trim().toLowerCase()
  return (
    normalized.startsWith('gs://') ||
    normalized.startsWith('https://storage.googleapis.com/') ||
    normalized.startsWith('http://storage.googleapis.com/') ||
    normalized.includes('.storage.googleapis.com/')
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
