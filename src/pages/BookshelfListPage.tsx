import { useDeferredValue, useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate } from 'react-router-dom'
import { bookshelfApi } from '../api/bookshelfApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, DownloadIcon } from '../components/icons'
import { useBookshelf } from '../hooks/useBookshelf'
import {
  defaultBookshelfListFilters,
  type BookshelfEntry,
  type BookshelfListFilters,
} from '../lib/bookshelfTypes'
import styles from '../styles/ResourcesListPage.module.css'

const PAGE_SIZE = 10

export function BookshelfListPage() {
  const navigate = useNavigate()
  const { items, pagination, summary, loading, error, fetch, remove } = useBookshelf(PAGE_SIZE)
  const [filters, setFilters] = useState<BookshelfListFilters>({
    ...defaultBookshelfListFilters,
    pageSize: PAGE_SIZE,
  })
  const [deleteCandidate, setDeleteCandidate] = useState<BookshelfEntry | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  const [activeActionId, setActiveActionId] = useState<string | null>(null)
  const deferredSearchTerm = useDeferredValue(filters.searchTerm)

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat('en-CA', {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [],
  )

  useEffect(() => {
    void fetch({
      ...filters,
      searchTerm: deferredSearchTerm,
    }).catch((fetchError) => {
      toast.error(fetchError instanceof Error ? fetchError.message : 'Could not load bookshelf entries.')
    })
  }, [deferredSearchTerm, fetch, filters.page, filters.pageSize, filters.searchTerm])

  function updateFilters(patch: Partial<BookshelfListFilters>) {
    setFilters((current) => ({
      ...current,
      ...patch,
    }))
  }

  async function handleDownloadBook(item: BookshelfEntry) {
    setActiveActionId(item.id)
    try {
      const blob = await bookshelfApi.getBookContent(item.id)
      const objectUrl = URL.createObjectURL(blob)
      triggerFileDownload(objectUrl, item.bookFileName || `${item.title}.download`)
      scheduleObjectUrlRevoke(objectUrl)
    } catch {
      toast.error('Could not download this book.')
    } finally {
      setActiveActionId(null)
    }
  }

  async function handleConfirmDelete() {
    if (!deleteCandidate) {
      return
    }

    setIsDeleting(true)
    try {
      await remove(deleteCandidate.id)
      const nextPage = pagination.page > 1 && items.length === 1 ? pagination.page - 1 : pagination.page

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
      toast.success('Book deleted.')
    } catch (deleteError) {
      toast.error(deleteError instanceof Error ? deleteError.message : 'Could not delete this book.')
    } finally {
      setIsDeleting(false)
    }
  }

  const totalItems = pagination.totalItems
  const rangeStart = totalItems > 0 ? (pagination.page - 1) * pagination.pageSize + 1 : 0
  const rangeEnd = totalItems > 0 ? Math.min(pagination.page * pagination.pageSize, totalItems) : 0

  return (
    <CmsAppShell activeKey="bookshelf">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: 'Bookshelf' }]} />

        <div className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>Bookshelf</h1>
            <p>
              Manage uploaded books, purchase links, author profiles, and optional artwork
              from one library view.
            </p>
          </div>
          <button
            type="button"
            className={styles.uploadButton}
            onClick={() => navigate('/bookshelf/new')}
          >
            <AddIcon size={16} />
            Add Book
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
              placeholder="Search by title, author, teaser, description, or file name..."
              aria-label="Search bookshelf"
              onChange={(event) => updateFilters({ searchTerm: event.target.value, page: 1 })}
            />
          </label>

          <div className={styles.typeFilter}>
            <span className={styles.filterIcon} aria-hidden="true">
              <BooksIcon />
            </span>
            <strong>{totalItems} books</strong>
          </div>
        </div>

        <div className={styles.categoryGrid}>
          <article className={styles.categoryCard}>
            <span className={styles.categoryIcon}>
              <CoverIcon />
            </span>
            <span className={styles.categoryTitle}>
              With Cover
              <span className={styles.categoryMeta}> ({summary.withCoverCount})</span>
            </span>
          </article>

          <article className={styles.categoryCard}>
            <span className={styles.categoryIcon}>
              <DocumentIcon />
            </span>
            <span className={styles.categoryTitle}>
              No Cover
              <span className={styles.categoryMeta}> ({summary.withoutCoverCount})</span>
            </span>
          </article>
        </div>

        <section className={styles.resultsPanel}>
          {error ? (
            <div className={styles.errorBox}>
              <p>{error}</p>
            </div>
          ) : null}

          {loading ? (
            <div className={styles.loadingWrap}>
              <Loader label="Loading books..." />
            </div>
          ) : items.length > 0 ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Title</th>
                      <th>Author</th>
                      <th>Updated</th>
                      <th>Book File</th>
                      <th>Media</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr key={item.id}>
                        <td>
                          <div className={styles.nameCell}>
                            <span className={styles.fileBadge}>{resolveFileBadge(item.bookFileName)}</span>
                            <div className={styles.nameStack}>
                              <span className={styles.fileTitle}>{item.title}</span>
                              <span>{buildTitleMeta(item)}</span>
                            </div>
                          </div>
                        </td>
                        <td>
                          <div className={styles.nameStack}>
                            <span className={styles.fileTitle}>{item.author}</span>
                            <span>{summarizeText(item.authorBio, 96)}</span>
                          </div>
                        </td>
                        <td>{formatDate(item.updatedAt, dateFormatter)}</td>
                        <td>{buildFileMeta(item.bookFileName, item.bookFileSize)}</td>
                        <td>
                          <span className={styles.categoryPill}>{resolveMediaStatus(item)}</span>
                        </td>
                        <td>
                          <div className={styles.actionsCell}>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={`Download ${item.title}`}
                              disabled={activeActionId === item.id}
                              onClick={() => void handleDownloadBook(item)}
                            >
                              <DownloadIcon size={16} />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label={`Edit ${item.title}`}
                              onClick={() => navigate(`/bookshelf/${item.id}/edit`)}
                            >
                              <PencilIcon />
                            </button>
                            <button
                              type="button"
                              className={`${styles.iconButton} ${styles.iconButtonDanger}`}
                              aria-label={`Delete ${item.title}`}
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
                  Showing {rangeStart} to {rangeEnd} of {totalItems} books
                </span>

                <div className={styles.pager}>
                  <button
                    type="button"
                    className={styles.pagerButton}
                    disabled={!pagination.hasPrev}
                    onClick={() => updateFilters({ page: pagination.page - 1 })}
                    aria-label="Previous page"
                  >
                    <ChevronLeftIcon />
                  </button>
                  <button
                    type="button"
                    className={styles.pagerButton}
                    disabled={!pagination.hasNext}
                    onClick={() => updateFilters({ page: pagination.page + 1 })}
                    aria-label="Next page"
                  >
                    <ChevronRightIcon />
                  </button>
                </div>
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>No books found</p>
              <p>Try a different search or add your first book to the bookshelf.</p>
            </div>
          )}
        </section>

        <ConfirmDialog
          open={Boolean(deleteCandidate)}
          title="Delete book?"
          body={`This will permanently delete "${deleteCandidate?.title ?? ''}" from the bookshelf. This action cannot be undone.`}
          confirmLabel="Delete"
          cancelLabel="Cancel"
          destructive
          busy={isDeleting}
          onConfirm={() => void handleConfirmDelete()}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function formatDate(value: string, formatter: Intl.DateTimeFormat) {
  const parsed = Date.parse(value)
  if (Number.isNaN(parsed)) {
    return value
  }
  return formatter.format(new Date(parsed))
}

function buildFileMeta(fileName: string, fileSize: number) {
  const parts = [fileName]
  if (fileSize > 0) {
    parts.push(formatFileSize(fileSize))
  }
  return parts.join(' | ')
}

function buildTitleMeta(item: BookshelfEntry) {
  const summary = item.bookTeaser.trim() || item.description.trim() || 'No teaser added yet.'
  return item.bookLink.trim() ? `${summary} | Purchase link ready` : summary
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

function resolveFileBadge(fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (!extension) {
    return 'BOOK'
  }
  return extension.length > 5 ? extension.slice(0, 5) : extension
}

function resolveMediaStatus(item: BookshelfEntry) {
  if (item.hasAuthorImage && item.hasCoverImage) {
    return 'Author + Cover'
  }
  if (item.hasAuthorImage) {
    return 'Author'
  }
  if (item.hasCoverImage) {
    return 'Cover'
  }
  return 'None'
}

function summarizeText(value: string, maxLength: number) {
  const normalized = value.trim()
  if (!normalized) {
    return 'Author bio not added yet.'
  }
  if (normalized.length <= maxLength) {
    return normalized
  }
  return `${normalized.slice(0, maxLength - 1).trimEnd()}...`
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

function BooksIcon() {
  return (
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
      <path d="M6 4.5h11v14H8.2A2.2 2.2 0 0 1 6 16.3V4.5Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
      <path d="M9 8h5M9 11h6M9 14h4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  )
}

function CoverIcon() {
  return (
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
      <rect x="4" y="5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="1.5" />
      <path d="M7 16l3.2-3.2 2.3 2.3 2.8-3.1L18 16" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      <circle cx="9" cy="9" r="1.2" fill="currentColor" />
    </svg>
  )
}

function DocumentIcon() {
  return (
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
      <path d="M7 3.5h7l3 3V20.5H7z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
      <path d="M14 3.5v3h3M9.5 11h5M9.5 14h5M9.5 17h3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
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
