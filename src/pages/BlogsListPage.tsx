import Dialog from '@mui/material/Dialog'
import DialogActions from '@mui/material/DialogActions'
import DialogContent from '@mui/material/DialogContent'
import DialogContentText from '@mui/material/DialogContentText'
import DialogTitle from '@mui/material/DialogTitle'
import DeleteOutlined from '@mui/icons-material/DeleteOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate } from 'react-router-dom'
import { blogsApi, type BlogListItem } from '../api/blogsApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { AddIcon, SearchIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import styles from '../styles/PagesListPage.module.css'

export function BlogsListPage() {
  const navigate = useNavigate()
  const [items, setItems] = useState<BlogListItem[]>([])
  const [status, setStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('loading')
  const [error, setError] = useState<string | null>(null)
  const [searchInput, setSearchInput] = useState('')
  const [searchTerm, setSearchTerm] = useState('')
  const [deleteCandidate, setDeleteCandidate] = useState<BlogListItem | null>(null)
  const [deletingBlogId, setDeletingBlogId] = useState<number | null>(null)

  useEffect(() => {
    const timer = window.setTimeout(() => setSearchTerm(searchInput), 250)
    return () => window.clearTimeout(timer)
  }, [searchInput])

  useEffect(() => {
    let cancelled = false

    async function loadBlogs() {
      setStatus('loading')
      setError(null)

      try {
        const response = await blogsApi.listBlogs({
          page: 1,
          pageSize: 100,
          searchTerm,
          sortBy: 'publish_date',
          sortOrder: 'desc',
        })
        if (cancelled) {
          return
        }
        setItems(response)
        setStatus('ready')
      } catch {
        if (cancelled) {
          return
        }
        setItems([])
        setStatus('error')
        setError('Could not load blogs right now.')
      }
    }

    void loadBlogs()

    return () => {
      cancelled = true
    }
  }, [searchTerm])

  useEffect(() => {
    if (deleteCandidate && !items.some((item) => item.id === deleteCandidate.id)) {
      setDeleteCandidate(null)
    }
  }, [deleteCandidate, items])

  const formatter = useMemo(
    () =>
      new Intl.DateTimeFormat(undefined, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [],
  )

  const isLoading = status === 'loading'

  async function handleDeleteBlog(item: BlogListItem) {
    setDeletingBlogId(item.id)
    try {
      await blogsApi.deleteBlog(item.id)
      setItems((current) => current.filter((candidate) => candidate.id !== item.id))
      setDeleteCandidate(null)
      toast.success('Blog deleted successfully.')
    } catch {
      toast.error('Could not delete this blog right now.')
    } finally {
      setDeletingBlogId(null)
    }
  }

  function formatLastModified(item: BlogListItem) {
    const editorName = item.updated_by_name?.trim() || 'Unknown editor'
    return (
      <>
        <div>{formatter.format(new Date(item.updated_at))}</div>
        <div className={styles.modifiedBy}>Updated by {editorName}</div>
      </>
    )
  }

  return (
    <CmsAppShell activeKey="blogs">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: 'Blogs' }]} />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>Blogs</h1>
            <p className={styles.subtitle}>
              Create and manage Living History Hub stories and blog modules.
            </p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/blogs/new')}
          >
            <AddIcon size={16} />
            Create blog
          </button>
        </header>

        <section className={styles.tableCard}>
          <div className={styles.tableToolbar}>
            <div className={styles.filterCardTop}>
              <div className={styles.searchField}>
                <span className={styles.searchIcon} aria-hidden="true">
                  <SearchIcon size={18} />
                </span>
                <input
                  type="search"
                  value={searchInput}
                  onChange={(event) => setSearchInput(event.target.value)}
                  placeholder="Search blogs..."
                  aria-label="Search blogs"
                />
              </div>

              <p className={styles.totalCount}>Total: {items.length} blogs</p>
            </div>
          </div>

          {error && <p className={styles.errorText}>{error}</p>}

          {isLoading && !items.length ? (
            <div className={styles.loaderWrap}>
              <Loader label="Loading blogs" />
            </div>
          ) : items.length ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Heading</th>
                      <th>Publish date</th>
                      <th>Description</th>
                      <th>Last updated</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr key={item.id}>
                        <td className={styles.pageTitleCell}>
                          <div className={styles.pageTitleStack}>
                            <span>{item.heading}</span>
                          </div>
                        </td>
                        <td>{formatter.format(new Date(item.publish_date))}</td>
                        <td>{item.description}</td>
                        <td>{formatLastModified(item)}</td>
                        <td>
                          <div className={styles.actions}>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label="View blog"
                              onClick={() => navigate(`/blogs/${item.id}`)}
                            >
                              <VisibilityOutlined fontSize="small" />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              aria-label="Edit blog"
                              onClick={() => navigate(`/blogs/${item.id}/edit`)}
                            >
                              <EditOutlined fontSize="small" />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label="Delete blog"
                              disabled={deletingBlogId === item.id}
                              onClick={() => setDeleteCandidate(item)}
                            >
                              <DeleteOutlined fontSize="small" />
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
                        <p className={styles.mobileCardTitle}>{item.heading}</p>
                        <div className={styles.mobileCardMeta}>
                          {formatter.format(new Date(item.publish_date))}
                        </div>
                      </div>
                    </div>
                    <div className={styles.mobileCardMeta}>{item.description}</div>
                    <div className={styles.mobileCardMeta}>{formatLastModified(item)}</div>
                    <div className={styles.actions}>
                      <button
                        type="button"
                        className={styles.iconButton}
                        aria-label="View blog"
                        onClick={() => navigate(`/blogs/${item.id}`)}
                      >
                        <VisibilityOutlined fontSize="small" />
                      </button>
                      <button
                        type="button"
                        className={styles.iconButton}
                        aria-label="Edit blog"
                        onClick={() => navigate(`/blogs/${item.id}/edit`)}
                      >
                        <EditOutlined fontSize="small" />
                      </button>
                      <button
                        type="button"
                        className={styles.iconButtonDanger}
                        aria-label="Delete blog"
                        disabled={deletingBlogId === item.id}
                        onClick={() => setDeleteCandidate(item)}
                      >
                        <DeleteOutlined fontSize="small" />
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>No blogs found</p>
              <p className={styles.emptyText}>
                Add your first Living History Hub story to get started.
              </p>
            </div>
          )}
        </section>

        <Dialog
          open={Boolean(deleteCandidate)}
          onClose={() => {
            if (!deleteCandidate || deletingBlogId !== deleteCandidate.id) {
              setDeleteCandidate(null)
            }
          }}
        >
          <DialogTitle>Delete blog</DialogTitle>
          <DialogContent>
            <DialogContentText>
              This will permanently delete "{deleteCandidate?.heading ?? ''}" from the
              blogs list. This action cannot be undone.
            </DialogContentText>
          </DialogContent>
          <DialogActions>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setDeleteCandidate(null)}
              disabled={Boolean(deleteCandidate && deletingBlogId === deleteCandidate.id)}
            >
              Cancel
            </button>
            <button
              type="button"
              className={styles.deleteButton}
              onClick={() => deleteCandidate && handleDeleteBlog(deleteCandidate)}
              disabled={Boolean(deleteCandidate && deletingBlogId === deleteCandidate.id)}
            >
              {deleteCandidate && deletingBlogId === deleteCandidate.id ? 'Deleting...' : 'Delete'}
            </button>
          </DialogActions>
        </Dialog>
      </div>
    </CmsAppShell>
  )
}
