import { useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate } from 'react-router-dom'
import { booksApi } from '../api/booksApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import styles from '../styles/BookCreatePage.module.css'

type BookFormState = {
  title: string
  description: string
}

export function BookCreatePage() {
  const navigate = useNavigate()
  const [form, setForm] = useState<BookFormState>({
    title: '',
    description: '',
  })
  const [isSaving, setIsSaving] = useState(false)

  async function handleCreateBook() {
    const title = form.title.trim()
    const description = form.description.trim()

    if (!title) {
      toast.error('Book title is required.')
      return
    }

    try {
      setIsSaving(true)
      const created = await booksApi.createBook({
        title,
        description,
        adminNotificationEmails: [],
      })
      toast.success('Book created.')
      navigate(`/books/${created.id}`)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to create the book.')
    } finally {
      setIsSaving(false)
    }
  }

  return (
    <CmsAppShell activeKey="books">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: 'Books', to: '/books' }, { label: 'New Book' }]} />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>Create Book</h1>
            <p className={styles.subtitle}>
              Start with the book title and description. The next page handles the one-time initial
              version setup, and every approved website request creates the next active version automatically.
            </p>
          </div>

          <div className={styles.headerActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/books')}>
              Back to Books
            </button>
            <button type="button" className={styles.primaryButton} onClick={() => void handleCreateBook()} disabled={isSaving}>
              {isSaving ? 'Creating...' : 'Create Book'}
            </button>
          </div>
        </header>

        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <div>
              <h2>Book Details</h2>
              <p>Add the book name and a short description for the CMS and website teams.</p>
            </div>
          </div>

          <div className={styles.fieldStack}>
            <label className={styles.field}>
              <span>Book Title</span>
              <input
                type="text"
                value={form.title}
                onChange={(event) => setForm((current) => ({ ...current, title: event.target.value }))}
                placeholder="Enter the book title"
              />
            </label>

            <label className={styles.field}>
              <span>Description</span>
              <textarea
                rows={5}
                value={form.description}
                onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))}
                placeholder="Add a short description"
              />
            </label>
          </div>
        </section>

        <section className={styles.infoCard}>
          <h2>How This Works</h2>
          <ul className={styles.infoList}>
            <li>Create the book record first, then complete the one-time initial version setup on the detail page.</li>
            <li>After that, every approved website request creates the next active version automatically.</li>
            <li>You can still switch the website back to any earlier version from the existing book detail page.</li>
          </ul>
        </section>
      </div>
    </CmsAppShell>
  )
}
