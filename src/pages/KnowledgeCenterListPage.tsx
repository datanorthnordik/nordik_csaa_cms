import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import { useDeferredValue, useEffect, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useSearchParams } from 'react-router-dom'
import {
  knowledgeCenterApi,
  type KnowledgeCenterListPageMeta,
  type KnowledgeCenterListSummary,
  type KnowledgeCenterSubmission,
  type KnowledgeCenterSubmissionStatus,
} from '../api/knowledgeCenterApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import {
  ChevronLeftIcon,
  ChevronRightIcon,
  SearchIcon,
} from '../components/icons'
import booksStyles from '../styles/BooksListPage.module.css'
import resourcesStyles from '../styles/ResourcesListPage.module.css'

const PAGE_SIZE = 10

const defaultPagination: KnowledgeCenterListPageMeta = {
  page: 1,
  pageSize: PAGE_SIZE,
  totalItems: 0,
  totalPages: 0,
  hasNext: false,
  hasPrev: false,
}

const defaultSummary: KnowledgeCenterListSummary = {
  openCount: 0,
  completedCount: 0,
}

export function KnowledgeCenterListPage() {
  const navigate = useNavigate()
  const [searchParams, setSearchParams] = useSearchParams()
  const activeView: KnowledgeCenterSubmissionStatus =
    searchParams.get('tab') === 'completed' ? 'completed' : 'open'

  const [items, setItems] = useState<KnowledgeCenterSubmission[]>([])
  const [pagination, setPagination] =
    useState<KnowledgeCenterListPageMeta>(defaultPagination)
  const [summary, setSummary] =
    useState<KnowledgeCenterListSummary>(defaultSummary)
  const [searchValue, setSearchValue] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const deferredSearchValue = useDeferredValue(searchValue)

  useEffect(() => {
    void loadSubmissions(activeView, pagination.page, deferredSearchValue)
  }, [activeView, deferredSearchValue, pagination.page])

  async function loadSubmissions(
    status: KnowledgeCenterSubmissionStatus,
    page: number,
    searchTerm: string,
  ) {
    try {
      setIsLoading(true)
      setError(null)

      const response = await knowledgeCenterApi.listSubmissions({
        page,
        pageSize: PAGE_SIZE,
        searchTerm,
        status,
      })

      setItems(response.items)
      setPagination(response.pagination)
      setSummary(response.summary)
    } catch (loadError) {
      const message =
        loadError instanceof Error
          ? loadError.message
          : 'Unable to load knowledge center requests.'
      setError(message)
      toast.error(message)
    } finally {
      setIsLoading(false)
    }
  }

  function switchTab(status: KnowledgeCenterSubmissionStatus) {
    const nextParams = new URLSearchParams(searchParams)
    if (status === 'completed') {
      nextParams.set('tab', 'completed')
    } else {
      nextParams.delete('tab')
    }
    setSearchParams(nextParams)
    setPagination((current) => ({
      ...current,
      page: 1,
    }))
  }

  function updatePage(page: number) {
    setPagination((current) => ({
      ...current,
      page,
    }))
  }

  const totalItems = pagination.totalItems
  const rangeStart =
    totalItems > 0 ? (pagination.page - 1) * pagination.pageSize + 1 : 0
  const rangeEnd =
    totalItems > 0
      ? Math.min(pagination.page * pagination.pageSize, pagination.totalItems)
      : 0

  return (
    <CmsAppShell activeKey="knowledgeCenter">
      <div className={booksStyles.page}>
        <Breadcrumb items={[{ label: 'Knowledge Center' }]} />

        <header className={booksStyles.pageHeader}>
          <div>
            <h1 className={booksStyles.title}>Knowledge Center</h1>
            <p className={booksStyles.subtitle}>
              Review Living History Hub contribution requests from the website,
              keep open work visible, and track everything that has already been
              completed.
            </p>
          </div>

          <div className={booksStyles.headerActions}>
            <button
              type="button"
              className={booksStyles.secondaryButton}
              onClick={() =>
                void loadSubmissions(
                  activeView,
                  pagination.page,
                  deferredSearchValue,
                )
              }
            >
              Refresh
            </button>
          </div>
        </header>

        <div
          className={booksStyles.tabBar}
          role="tablist"
          aria-label="Knowledge center views"
        >
          <button
            type="button"
            role="tab"
            aria-selected={activeView === 'open'}
            className={[
              booksStyles.tabButton,
              activeView === 'open' ? booksStyles.tabButtonActive : '',
            ].join(' ')}
            onClick={() => switchTab('open')}
          >
            Open Requests
            <span className={booksStyles.tabCount}>{summary.openCount}</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeView === 'completed'}
            className={[
              booksStyles.tabButton,
              activeView === 'completed' ? booksStyles.tabButtonActive : '',
            ].join(' ')}
            onClick={() => switchTab('completed')}
          >
            Completed Requests
            <span className={booksStyles.tabCount}>
              {summary.completedCount}
            </span>
          </button>
        </div>

        <section className={booksStyles.tableCard}>
          <div className={booksStyles.tableToolbar}>
            <div className={booksStyles.filterCardTop}>
              <label className={booksStyles.searchField}>
                <span className={booksStyles.searchIcon} aria-hidden="true">
                  <SearchIcon size={18} />
                </span>
                <input
                  type="search"
                  value={searchValue}
                  onChange={(event) => {
                    setSearchValue(event.target.value)
                    setPagination((current) => ({
                      ...current,
                      page: 1,
                    }))
                  }}
                  placeholder={
                    activeView === 'open'
                      ? 'Search open requests'
                      : 'Search completed requests'
                  }
                  aria-label={
                    activeView === 'open'
                      ? 'Search open requests'
                      : 'Search completed requests'
                  }
                />
              </label>

              <p className={booksStyles.totalCount}>
                {pagination.totalItems}{' '}
                {activeView === 'open' ? 'open' : 'completed'} request
                {pagination.totalItems === 1 ? '' : 's'}
              </p>
            </div>
          </div>

          {isLoading ? (
            <div className={booksStyles.loaderWrap}>
              <Loader />
            </div>
          ) : error ? (
            <div className={booksStyles.emptyState}>
              <h2>Unable to load requests</h2>
              <p>{error}</p>
            </div>
          ) : items.length === 0 ? (
            <div className={booksStyles.emptyState}>
              <h2>
                {activeView === 'open'
                  ? 'No open requests'
                  : 'No completed requests'}
              </h2>
              <p>
                {activeView === 'open'
                  ? 'New Living History Hub submissions will appear here when people share their stories.'
                  : 'Completed Living History Hub submissions will appear here after your team finishes them.'}
              </p>
            </div>
          ) : (
            <>
              <div className={booksStyles.tableWrap}>
                <table className={booksStyles.table}>
                  <thead>
                    <tr>
                      <th>Request</th>
                      <th>Submitter</th>
                      <th>Type</th>
                      {activeView === 'open' ? (
                        <>
                          <th>Contact</th>
                          <th>Submitted</th>
                        </>
                      ) : (
                        <>
                          <th>Completed By</th>
                          <th>Completed</th>
                        </>
                      )}
                      <th>Open</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr
                        key={item.id}
                        className={booksStyles.row}
                        onClick={() => navigate(`/knowledge-center/${item.id}`)}
                      >
                        <td>
                          <span className={booksStyles.metaPill}>
                            Request #{item.id}
                          </span>
                        </td>
                        <td className={booksStyles.primaryCell}>
                          <div className={booksStyles.primaryStack}>
                            <strong>{item.submitterName}</strong>
                          </div>
                        </td>
                        <td>{formatSubmissionType(item.submissionType)}</td>
                        {activeView === 'open' ? (
                          <>
                            <td>
                              <div className={booksStyles.primaryStack}>
                                <strong>{item.submitterEmail}</strong>
                                <span>
                                  {item.submitterPhone || 'No phone provided'}
                                </span>
                              </div>
                            </td>
                            <td>{formatDate(item.createdAt)}</td>
                          </>
                        ) : (
                          <>
                            <td>
                              <div className={booksStyles.primaryStack}>
                                <strong>
                                  {item.completedBy?.name || 'Unknown reviewer'}
                                </strong>
                                <span>
                                  {item.completedBy?.email || 'No email stored'}
                                </span>
                              </div>
                            </td>
                            <td>{formatDate(item.completedAt ?? item.updatedAt)}</td>
                          </>
                        )}
                        <td>
                          <button
                            type="button"
                            className={booksStyles.iconButton}
                            aria-label={`Open request ${item.id}`}
                            onClick={(event) => {
                              event.stopPropagation()
                              navigate(`/knowledge-center/${item.id}`)
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

              <div className={booksStyles.cardList}>
                {items.map((item) => (
                  <article key={item.id} className={booksStyles.mobileCard}>
                    <div className={booksStyles.mobileCardTop}>
                      <div>
                        <p className={booksStyles.mobileCardTitle}>
                          {item.submitterName}
                        </p>
                      </div>
                      <span className={booksStyles.metaPill}>
                        {formatSubmissionType(item.submissionType)}
                      </span>
                    </div>

                    <div className={booksStyles.mobileMeta}>
                      <span>{item.submitterEmail}</span>
                      <span>
                        {activeView === 'open'
                          ? `Submitted ${formatDate(item.createdAt)}`
                          : `Completed ${formatDate(item.completedAt ?? item.updatedAt)}`}
                      </span>
                      <span>
                        {activeView === 'open'
                          ? item.submitterPhone || 'No phone provided'
                          : item.completedBy?.name || 'Unknown reviewer'}
                      </span>
                    </div>

                    <div className={booksStyles.mobileActions}>
                      <button
                        type="button"
                        className={booksStyles.secondaryButton}
                        onClick={() => navigate(`/knowledge-center/${item.id}`)}
                      >
                        Review Request
                      </button>
                    </div>
                  </article>
                ))}
              </div>

              <div className={resourcesStyles.footer}>
                <span className={resourcesStyles.footerInfo}>
                  Showing {rangeStart} to {rangeEnd} of {totalItems} requests
                </span>

                <div className={resourcesStyles.pager}>
                  <button
                    type="button"
                    className={resourcesStyles.pagerButton}
                    disabled={!pagination.hasPrev}
                    onClick={() => updatePage(pagination.page - 1)}
                    aria-label="Previous page"
                  >
                    <ChevronLeftIcon size={16} />
                  </button>
                  <button
                    type="button"
                    className={resourcesStyles.pagerButton}
                    disabled={!pagination.hasNext}
                    onClick={() => updatePage(pagination.page + 1)}
                    aria-label="Next page"
                  >
                    <ChevronRightIcon size={16} />
                  </button>
                </div>
              </div>
            </>
          )}
        </section>
      </div>
    </CmsAppShell>
  )
}

function formatSubmissionType(value: KnowledgeCenterSubmission['submissionType']) {
  switch (value) {
    case 'post':
      return 'Story'
    case 'video':
      return 'Video'
    case 'both':
      return 'Story + Video'
    default:
      return value
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
