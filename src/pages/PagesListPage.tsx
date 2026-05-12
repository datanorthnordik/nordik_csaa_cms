import Dialog from '@mui/material/Dialog'
import DialogActions from '@mui/material/DialogActions'
import DialogContent from '@mui/material/DialogContent'
import DialogContentText from '@mui/material/DialogContentText'
import DialogTitle from '@mui/material/DialogTitle'
import {
  DeleteOutline,
  EditOutlined,
  VisibilityOutlined,
} from '@mui/icons-material'
import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import {
  AddIcon,
  ChevronLeftIcon,
  ChevronRightIcon,
  SearchIcon,
} from '../components/icons'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import type { PageListItem, PageStatusFilter } from '../api/pagesApi'
import {
  deletePage as deletePageAction,
  fetchPages,
  selectPageList,
  selectPageListFilters,
  setPageListFilters,
} from '../store/pagesSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import styles from '../styles/PagesListPage.module.css'

type PaginationToken = number | 'ellipsis'

export function PagesListPage() {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const filters = useAppSelector(selectPageListFilters)
  const { items, pagination, status, error } = useAppSelector(selectPageList)
  const [searchInput, setSearchInput] = useState(filters.searchTerm)
  const [deleteCandidate, setDeleteCandidate] = useState<PageListItem | null>(null)
  const [deletingPageId, setDeletingPageId] = useState<number | null>(null)

  useEffect(() => {
    setSearchInput(filters.searchTerm)
  }, [filters.searchTerm])

  useEffect(() => {
    void dispatch(fetchPages(filters))
  }, [dispatch, filters])

  useEffect(() => {
    const timer = window.setTimeout(() => {
      if (searchInput !== filters.searchTerm) {
        dispatch(
          setPageListFilters({
            searchTerm: searchInput,
            page: 1,
          }),
        )
      }
    }, 250)

    return () => window.clearTimeout(timer)
  }, [dispatch, filters.searchTerm, searchInput])

  useEffect(() => {
    if (deleteCandidate && !items.some((item) => item.id === deleteCandidate.id)) {
      setDeleteCandidate(null)
    }
  }, [deleteCandidate, items])

  const formatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  const isLoading = status === 'loading'
  const totalItems = pagination?.total_items ?? 0
  const rangeStart =
    pagination && totalItems > 0 ? (pagination.page - 1) * pagination.page_size + 1 : 0
  const rangeEnd =
    pagination && totalItems > 0
      ? Math.min(pagination.page * pagination.page_size, totalItems)
      : 0

  function changePage(nextPage: number) {
    if (!pagination || nextPage < 1 || nextPage > pagination.total_pages) {
      return
    }

    dispatch(setPageListFilters({ page: nextPage }))
  }

  function changeStatus(statusValue: PageStatusFilter) {
    dispatch(
      setPageListFilters({
        status: statusValue,
        page: 1,
      }),
    )
  }

  async function handleDeletePage(item: PageListItem) {
    setDeletingPageId(item.id)

    try {
      await dispatch(deletePageAction(item.id)).unwrap()
      setDeleteCandidate(null)

      const nextPage =
        pagination && pagination.page > 1 && items.length === 1
          ? pagination.page - 1
          : filters.page

      dispatch(setPageListFilters({ page: nextPage }))
      toast.success(t('pages.feedback.deleted'))
    } catch {
      return
    } finally {
      setDeletingPageId(null)
    }
  }

  function formatLastModified(item: PageListItem) {
    const name = item.modified_by_name?.trim() || t('pages.list.unknownEditor')
    return (
      <>
        <div>{formatter.format(new Date(item.last_modified))}</div>
        <div className={styles.modifiedBy}>{t('pages.list.modifiedBy', { name })}</div>
      </>
    )
  }

  function buildPaginationTokens(currentPage: number, totalPages: number): PaginationToken[] {
    if (totalPages <= 5) {
      return Array.from({ length: totalPages }, (_, index) => index + 1)
    }

    if (currentPage <= 3) {
      return [1, 2, 3, 'ellipsis', totalPages]
    }

    if (currentPage >= totalPages - 2) {
      return [1, 'ellipsis', totalPages - 2, totalPages - 1, totalPages]
    }

    return [1, 'ellipsis', currentPage, 'ellipsis', totalPages]
  }

  const paginationTokens = pagination
    ? buildPaginationTokens(pagination.page, pagination.total_pages)
    : []

  return (
    <CmsAppShell activeKey="pages">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('pages.breadcrumb.pages') }]} />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('pages.list.title')}</h1>
            <p className={styles.subtitle}>{t('pages.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/pages/new')}
          >
            <AddIcon size={16} />
            {t('pages.list.create')}
          </button>
        </header>

        <section className={styles.filterCard}>
          <div className={styles.searchField}>
            <span className={styles.searchIcon} aria-hidden="true">
              <SearchIcon size={18} />
            </span>
            <input
              type="search"
              value={searchInput}
              onChange={(event) => setSearchInput(event.target.value)}
              placeholder={t('pages.filters.searchPlaceholder')}
              aria-label={t('pages.filters.searchPlaceholder')}
            />
          </div>

          <label className={styles.filterSelect}>
            <span className={styles.visuallyHidden}>{t('pages.filters.status')}</span>
            <select
              value={filters.status}
              onChange={(event) => changeStatus(event.target.value as PageStatusFilter)}
            >
              <option value="">{t('pages.filters.allStatuses')}</option>
              <option value="draft">{t('pages.status.draft')}</option>
              <option value="published">{t('pages.status.live')}</option>
            </select>
          </label>

          <p className={styles.totalCount}>
            {t('pages.list.total', { count: totalItems })}
          </p>
        </section>

        <section className={styles.tableCard}>
          {error && <p className={styles.errorText}>{error}</p>}

          {isLoading && !items.length ? (
            <div className={styles.loaderWrap}>
              <Loader label={t('pages.common.loading')} />
            </div>
          ) : items.length ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>{t('pages.list.columns.pageTitle')}</th>
                      <th>{t('pages.list.columns.urlSlug')}</th>
                      <th>{t('pages.list.columns.status')}</th>
                      <th>{t('pages.list.columns.lastModified')}</th>
                      <th>{t('pages.list.columns.actions')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr key={item.id}>
                        <td className={styles.pageTitleCell}>{item.page_title}</td>
                        <td>
                          <span className={styles.slugPill}>{item.url_slug}</span>
                        </td>
                        <td>
                          <span
                            className={[
                              styles.statusBadge,
                              item.status === 'published'
                                ? styles.statusPublished
                                : styles.statusDraft,
                            ].join(' ')}
                          >
                            {item.status === 'published'
                              ? t('pages.status.live')
                              : t('pages.status.draft')}
                          </span>
                        </td>
                        <td>{formatLastModified(item)}</td>
                        <td>
                          <div className={styles.actions}>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={t('pages.list.view')}
                              onClick={() => navigate(`/pages/${item.id}`)}
                            >
                              <VisibilityOutlined fontSize="small" />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={t('pages.list.edit')}
                              onClick={() => navigate(`/pages/${item.id}/edit`)}
                            >
                              <EditOutlined fontSize="small" />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label={t('pages.list.delete')}
                              disabled={deletingPageId === item.id}
                              onClick={() => setDeleteCandidate(item)}
                            >
                              <DeleteOutline fontSize="small" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className={styles.cardList}>
                {items.map((item) => (
                  <article key={item.id} className={styles.mobileCard}>
                    <div className={styles.mobileCardTop}>
                      <div>
                        <p className={styles.mobileCardTitle}>{item.page_title}</p>
                        <span className={styles.slugPill}>{item.url_slug}</span>
                      </div>
                      <span
                        className={[
                          styles.statusBadge,
                          item.status === 'published'
                            ? styles.statusPublished
                            : styles.statusDraft,
                        ].join(' ')}
                      >
                        {item.status === 'published'
                          ? t('pages.status.live')
                          : t('pages.status.draft')}
                      </span>
                    </div>
                    <div className={styles.mobileCardMeta}>{formatLastModified(item)}</div>
                    <div className={styles.actions}>
                      <button
                        type="button"
                        className={styles.iconButton}
                        aria-label={t('pages.list.view')}
                        onClick={() => navigate(`/pages/${item.id}`)}
                      >
                        <VisibilityOutlined fontSize="small" />
                      </button>
                      <button
                        type="button"
                        className={styles.iconButton}
                        aria-label={t('pages.list.edit')}
                        onClick={() => navigate(`/pages/${item.id}/edit`)}
                      >
                        <EditOutlined fontSize="small" />
                      </button>
                      <button
                        type="button"
                        className={styles.iconButtonDanger}
                        aria-label={t('pages.list.delete')}
                        disabled={deletingPageId === item.id}
                        onClick={() => setDeleteCandidate(item)}
                      >
                        <DeleteOutline fontSize="small" />
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>{t('pages.list.emptyTitle')}</p>
              <p className={styles.emptyText}>{t('pages.list.emptyText')}</p>
            </div>
          )}

          {pagination && pagination.total_pages > 1 && (
            <footer className={styles.pagination}>
              <p className={styles.paginationSummary}>
                {t('pages.list.resultsLabel', {
                  start: rangeStart,
                  end: rangeEnd,
                  total: totalItems,
                })}
              </p>

              <div className={styles.paginationControls}>
                <button
                  type="button"
                  className={styles.paginationButton}
                  disabled={!pagination.has_prev}
                  onClick={() => changePage(pagination.page - 1)}
                >
                  <ChevronLeftIcon size={18} />
                </button>

                {paginationTokens.map((token, index) =>
                  token === 'ellipsis' ? (
                    <span key={`ellipsis-${index}`} className={styles.ellipsis}>
                      …
                    </span>
                  ) : (
                    <button
                      key={token}
                      type="button"
                      className={[
                        styles.pageNumberButton,
                        token === pagination.page ? styles.pageNumberActive : '',
                      ].join(' ')}
                      onClick={() => changePage(token)}
                    >
                      {token}
                    </button>
                  ),
                )}

                <button
                  type="button"
                  className={styles.paginationButton}
                  disabled={!pagination.has_next}
                  onClick={() => changePage(pagination.page + 1)}
                >
                  <ChevronRightIcon size={18} />
                </button>
              </div>
            </footer>
          )}
        </section>

        <Dialog
          open={Boolean(deleteCandidate)}
          onClose={() => {
            if (!deleteCandidate || deletingPageId !== deleteCandidate.id) {
              setDeleteCandidate(null)
            }
          }}
        >
          <DialogTitle>{t('pages.list.deleteDialogTitle')}</DialogTitle>
          <DialogContent>
            <DialogContentText>
              {t('pages.list.deleteDialogDescription', {
                title: deleteCandidate?.page_title ?? '',
              })}
            </DialogContentText>
          </DialogContent>
          <DialogActions>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setDeleteCandidate(null)}
              disabled={Boolean(deleteCandidate && deletingPageId === deleteCandidate.id)}
            >
              {t('pages.list.cancelDelete')}
            </button>
            <button
              type="button"
              className={styles.deleteButton}
              onClick={() => deleteCandidate && handleDeletePage(deleteCandidate)}
              disabled={Boolean(deleteCandidate && deletingPageId === deleteCandidate.id)}
            >
              {deleteCandidate && deletingPageId === deleteCandidate.id
                ? t('pages.common.loading')
                : t('pages.list.delete')}
            </button>
          </DialogActions>
        </Dialog>
      </div>
    </CmsAppShell>
  )
}
