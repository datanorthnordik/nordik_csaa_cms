import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import { useDeferredValue, useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useSearchParams } from 'react-router-dom'
import { type BookSubmission, type BookSummary, booksApi } from '../api/booksApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { AddIcon, SearchIcon } from '../components/icons'
import styles from '../styles/BooksListPage.module.css'

type BooksView = 'books' | 'requests'

export type BookRequestRow = {
  submissionId: number
  bookId: number
  bookTitle: string
  bookDescription: string
  bookVersionId: number
  bookVersionNumber: number
  requestedSection: string
  submitterEmail: string
  createdAt: string
}

export function buildPendingRequestRows(books: BookSummary[], submissionsByBook: Map<number, BookSubmission[]>) {
  return books
    .flatMap((book) =>
      (submissionsByBook.get(book.id) ?? []).map(
        (submission): BookRequestRow => ({
          submissionId: submission.id,
          bookId: book.id,
          bookTitle: book.title,
          bookDescription: book.description,
          bookVersionId: submission.bookVersionId,
          bookVersionNumber: submission.bookVersionNumber,
          requestedSection: submission.targetSectionName || submission.newSectionName || 'No section selected',
          submitterEmail: submission.submitterEmail,
          createdAt: submission.createdAt,
        }),
      ),
    )
    .sort((left, right) => {
      const dateDiff = new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime()
      if (dateDiff !== 0) {
        return dateDiff
      }
      return right.submissionId - left.submissionId
    })
}

export function filterPendingRequestRows(rows: BookRequestRow[], searchValue: string) {
  const needle = searchValue.trim().toLowerCase()
  if (!needle) {
    return rows
  }

  return rows.filter((row) =>
    [
      row.bookTitle,
      row.bookDescription,
      row.requestedSection,
      row.submitterEmail,
      `version ${row.bookVersionNumber}`,
      `request ${row.submissionId}`,
    ]
      .join(' ')
      .toLowerCase()
      .includes(needle),
  )
}

export function BooksListPage() {
  const navigate = useNavigate()
  const [searchParams, setSearchParams] = useSearchParams()
  const activeView: BooksView = searchParams.get('tab') === 'requests' ? 'requests' : 'books'

  const [books, setBooks] = useState<BookSummary[]>([])
  const [pendingRequests, setPendingRequests] = useState<BookRequestRow[]>([])
  const [searchValue, setSearchValue] = useState('')
  const [isBooksLoading, setIsBooksLoading] = useState(true)
  const [isRequestsLoading, setIsRequestsLoading] = useState(false)
  const [hasLoadedRequests, setHasLoadedRequests] = useState(false)
  const deferredSearchValue = useDeferredValue(searchValue)

  useEffect(() => {
    void loadBooks()
  }, [])

  useEffect(() => {
    if (activeView === 'requests' && !hasLoadedRequests && !isBooksLoading) {
      void loadPendingRequests(books)
    }
  }, [activeView, books, hasLoadedRequests, isBooksLoading])

  async function loadBooks() {
    try {
      setIsBooksLoading(true)
      const items = await booksApi.listBooks()
      setBooks(items)
      return items
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load books.')
      return []
    } finally {
      setIsBooksLoading(false)
    }
  }

  async function loadPendingRequests(sourceBooks = books) {
    try {
      setIsRequestsLoading(true)

      const booksToUse = sourceBooks.length > 0 ? sourceBooks : await booksApi.listBooks()
      if (sourceBooks.length === 0) {
        setBooks(booksToUse)
      }

      const pendingBooks = booksToUse.filter((book) => book.pendingSubmissionCount > 0)
      if (pendingBooks.length === 0) {
        setPendingRequests([])
        setHasLoadedRequests(true)
        return
      }

      const groups = await Promise.all(
        pendingBooks.map(async (book) => ({
          bookId: book.id,
          submissions: await booksApi.listSubmissions(book.id, 0, 'pending'),
        })),
      )

      const submissionsByBook = new Map(groups.map((group) => [group.bookId, group.submissions]))
      setPendingRequests(buildPendingRequestRows(booksToUse, submissionsByBook))
      setHasLoadedRequests(true)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load pending requests.')
    } finally {
      setIsRequestsLoading(false)
    }
  }

  async function handleRefresh() {
    const refreshedBooks = await loadBooks()
    if (activeView === 'requests') {
      await loadPendingRequests(refreshedBooks)
    }
  }

  const filteredBooks = useMemo(() => {
    const needle = deferredSearchValue.trim().toLowerCase()
    return books
      .filter((book) => {
        if (!needle) {
          return true
        }

        return [book.title, book.description].join(' ').toLowerCase().includes(needle)
      })
      .sort((left, right) => {
        const updatedDiff = new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime()
        if (updatedDiff !== 0) {
          return updatedDiff
        }
        return left.title.localeCompare(right.title)
      })
  }, [books, deferredSearchValue])

  const filteredRequests = useMemo(
    () => filterPendingRequestRows(pendingRequests, deferredSearchValue),
    [pendingRequests, deferredSearchValue],
  )

  return (
    <CmsAppShell activeKey="books">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: 'Books' }]} />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>Books</h1>
            <p className={styles.subtitle}>
              {activeView === 'books'
                ? 'Open a book to create its initial version once, inspect later approval versions, and switch the active version when needed.'
                : 'Review pending website requests in a dedicated queue and open each request in its own approval page.'}
            </p>
          </div>

          <div className={styles.headerActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => void handleRefresh()}>
              Refresh
            </button>
            <button type="button" className={styles.primaryButton} onClick={() => navigate('/books/new')}>
              <AddIcon size={16} />
              New Book
            </button>
          </div>
        </header>

        <div className={styles.tabBar} role="tablist" aria-label="Books views">
          <button
            type="button"
            role="tab"
            aria-selected={activeView === 'books'}
            className={[styles.tabButton, activeView === 'books' ? styles.tabButtonActive : ''].join(' ')}
            onClick={() => setSearchParams({})}
          >
            Books
            <span className={styles.tabCount}>{books.length}</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeView === 'requests'}
            className={[styles.tabButton, activeView === 'requests' ? styles.tabButtonActive : ''].join(' ')}
            onClick={() => setSearchParams({ tab: 'requests' })}
          >
            Open Requests
            <span className={styles.tabCount}>{pendingRequests.length}</span>
          </button>
        </div>

        <section className={styles.tableCard}>
          <div className={styles.tableToolbar}>
            <div className={styles.filterCardTop}>
              <label className={styles.searchField}>
                <span className={styles.searchIcon} aria-hidden="true">
                  <SearchIcon size={18} />
                </span>
                <input
                  type="search"
                  value={searchValue}
                  onChange={(event) => setSearchValue(event.target.value)}
                  placeholder={activeView === 'books' ? 'Search books' : 'Search requests'}
                  aria-label={activeView === 'books' ? 'Search books' : 'Search requests'}
                />
              </label>

              <p className={styles.totalCount}>
                {activeView === 'books'
                  ? `${filteredBooks.length} book${filteredBooks.length === 1 ? '' : 's'}`
                  : `${filteredRequests.length} open request${filteredRequests.length === 1 ? '' : 's'}`}
              </p>
            </div>
          </div>

          {activeView === 'books' ? renderBooksTable() : renderRequestsTable()}
        </section>
      </div>
    </CmsAppShell>
  )

  function renderBooksTable() {
    if (isBooksLoading) {
      return (
        <div className={styles.loaderWrap}>
          <Loader />
        </div>
      )
    }

    if (filteredBooks.length === 0) {
      return (
        <div className={styles.emptyState}>
          <h2>No books found</h2>
          <p>Try a different search or create a new book.</p>
        </div>
      )
    }

    return (
      <>
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Book</th>
                <th>Pending Requests</th>
                <th>Active Version</th>
                <th>Versions</th>
                <th>Updated</th>
                <th>Open</th>
              </tr>
            </thead>
            <tbody>
              {filteredBooks.map((book) => (
                <tr key={book.id} className={styles.row} onClick={() => navigate(`/books/${book.id}`)}>
                  <td className={styles.primaryCell}>
                    <div className={styles.primaryStack}>
                      <strong>{book.title}</strong>
                      <span>{book.description || 'No description added yet.'}</span>
                    </div>
                  </td>
                  <td>
                    <span className={book.pendingSubmissionCount > 0 ? styles.badgeAlert : styles.badgeMuted}>
                      {book.pendingSubmissionCount}
                    </span>
                  </td>
                  <td>
                    <span className={styles.metaPill}>
                      {book.activeVersionNumber ? `Version ${book.activeVersionNumber}` : 'None'}
                    </span>
                  </td>
                  <td>{book.versionCount}</td>
                  <td>{formatDate(book.updatedAt)}</td>
                  <td>
                    <button
                      type="button"
                      className={styles.iconButton}
                      aria-label={`Open ${book.title}`}
                      onClick={(event) => {
                        event.stopPropagation()
                        navigate(`/books/${book.id}`)
                      }}
                    >
                      <VisibilityOutlined fontSize="small" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className={styles.cardList}>
          {filteredBooks.map((book) => (
            <article key={book.id} className={styles.mobileCard}>
              <div className={styles.mobileCardTop}>
                <div>
                  <p className={styles.mobileCardTitle}>{book.title}</p>
                  <p className={styles.mobileCardText}>{book.description || 'No description added yet.'}</p>
                </div>
                <span className={book.pendingSubmissionCount > 0 ? styles.badgeAlert : styles.badgeMuted}>
                  {book.pendingSubmissionCount} pending
                </span>
              </div>

              <div className={styles.mobileMeta}>
                <span>{book.activeVersionNumber ? `Active version ${book.activeVersionNumber}` : 'No active version'}</span>
                <span>{book.versionCount} version{book.versionCount === 1 ? '' : 's'}</span>
                <span>Updated {formatDate(book.updatedAt)}</span>
              </div>

              <div className={styles.mobileActions}>
                <button type="button" className={styles.secondaryButton} onClick={() => navigate(`/books/${book.id}`)}>
                  Open Book
                </button>
              </div>
            </article>
          ))}
        </div>
      </>
    )
  }

  function renderRequestsTable() {
    if (isRequestsLoading) {
      return (
        <div className={styles.loaderWrap}>
          <Loader />
        </div>
      )
    }

    if (filteredRequests.length === 0) {
      return (
        <div className={styles.emptyState}>
          <h2>No open requests</h2>
          <p>Pending website requests will appear here when they need review.</p>
        </div>
      )
    }

    return (
      <>
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Request</th>
                <th>Book</th>
                <th>Version</th>
                <th>Requested Section</th>
                <th>Submitter</th>
                <th>Submitted</th>
                <th>Open</th>
              </tr>
            </thead>
            <tbody>
              {filteredRequests.map((request) => (
                <tr
                  key={`${request.bookId}-${request.submissionId}`}
                  className={styles.row}
                  onClick={() => navigate(`/books/requests/${request.bookId}/${request.submissionId}`)}
                >
                  <td>
                    <span className={styles.metaPill}>Request #{request.submissionId}</span>
                  </td>
                  <td className={styles.primaryCell}>
                    <div className={styles.primaryStack}>
                      <strong>{request.bookTitle}</strong>
                      <span>{request.bookDescription || 'No description added yet.'}</span>
                    </div>
                  </td>
                  <td>Version {request.bookVersionNumber}</td>
                  <td>{request.requestedSection}</td>
                  <td>{request.submitterEmail || 'Not provided'}</td>
                  <td>{formatDate(request.createdAt)}</td>
                  <td>
                    <button
                      type="button"
                      className={styles.iconButton}
                      aria-label={`Open request ${request.submissionId}`}
                      onClick={(event) => {
                        event.stopPropagation()
                        navigate(`/books/requests/${request.bookId}/${request.submissionId}`)
                      }}
                    >
                      <VisibilityOutlined fontSize="small" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className={styles.cardList}>
          {filteredRequests.map((request) => (
            <article key={`${request.bookId}-${request.submissionId}`} className={styles.mobileCard}>
              <div className={styles.mobileCardTop}>
                <div>
                  <p className={styles.mobileCardTitle}>Request #{request.submissionId}</p>
                  <p className={styles.mobileCardText}>{request.bookTitle}</p>
                </div>
                <span className={styles.metaPill}>Version {request.bookVersionNumber}</span>
              </div>

              <div className={styles.mobileMeta}>
                <span>{request.requestedSection}</span>
                <span>{request.submitterEmail || 'Submitter email not provided'}</span>
                <span>Submitted {formatDate(request.createdAt)}</span>
              </div>

              <div className={styles.mobileActions}>
                <button
                  type="button"
                  className={styles.secondaryButton}
                  onClick={() => navigate(`/books/requests/${request.bookId}/${request.submissionId}`)}
                >
                  Review Request
                </button>
              </div>
            </article>
          ))}
        </div>
      </>
    )
  }
}

function formatDate(value: string) {
  const parsed = Date.parse(value)
  if (Number.isNaN(parsed)) {
    return value
  }

  return new Intl.DateTimeFormat('en-CA', {
    month: 'short',
    day: '2-digit',
    year: 'numeric',
  }).format(new Date(parsed))
}
