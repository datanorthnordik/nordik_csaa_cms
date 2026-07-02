import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useParams } from 'react-router-dom'
import { bookshelfApi } from '../api/bookshelfApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { useBookshelf } from '../hooks/useBookshelf'
import {
  BOOKSHELF_AUTHOR_IMAGE_ACCEPT,
  BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB,
  BOOKSHELF_BOOK_ACCEPT,
  BOOKSHELF_BOOK_MAX_FILE_SIZE_MB,
  BOOKSHELF_COVER_ACCEPT,
  BOOKSHELF_COVER_MAX_FILE_SIZE_MB,
  validateBookshelfAuthorImageFile,
  validateBookshelfBookFile,
  validateBookshelfCoverFile,
} from '../lib/bookshelfUpload'
import type {
  BookshelfEntry,
  BookshelfFormErrors,
  BookshelfFormState,
} from '../lib/bookshelfTypes'
import styles from '../styles/ResourceEditorPage.module.css'

type BookshelfEditorPageProps = {
  mode?: 'create' | 'edit'
}

type BookshelfPreviewKind = 'book' | 'authorImage' | 'cover'

function emptyFormState(): BookshelfFormState {
  return {
    author: '',
    title: '',
    bookLink: '',
    authorBio: '',
    bookTeaser: '',
    description: '',
  }
}

function bookToFormState(book: BookshelfEntry): BookshelfFormState {
  return {
    author: book.author,
    title: book.title,
    bookLink: book.bookLink,
    authorBio: book.authorBio,
    bookTeaser: book.bookTeaser,
    description: book.description,
  }
}

export function BookshelfEditorPage({ mode = 'create' }: BookshelfEditorPageProps) {
  const navigate = useNavigate()
  const { id } = useParams()
  const { create, update, remove } = useBookshelf()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentBook, setCurrentBook] = useState<BookshelfEntry | null>(null)
  const [isLoading, setIsLoading] = useState(isEditMode)
  const [form, setForm] = useState<BookshelfFormState>(emptyFormState())
  const [selectedBookFile, setSelectedBookFile] = useState<File | null>(null)
  const [selectedAuthorImageFile, setSelectedAuthorImageFile] = useState<File | null>(null)
  const [selectedCoverFile, setSelectedCoverFile] = useState<File | null>(null)
  const [removeAuthorImage, setRemoveAuthorImage] = useState(false)
  const [removeCoverImage, setRemoveCoverImage] = useState(false)
  const [bookPreviewUrl, setBookPreviewUrl] = useState<string | null>(null)
  const [authorImagePreviewUrl, setAuthorImagePreviewUrl] = useState<string | null>(null)
  const [coverPreviewUrl, setCoverPreviewUrl] = useState<string | null>(null)
  const bookPreviewUrlRef = useRef<string | null>(null)
  const authorImagePreviewUrlRef = useRef<string | null>(null)
  const coverPreviewUrlRef = useRef<string | null>(null)
  const [errors, setErrors] = useState<BookshelfFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const [activeBookAction, setActiveBookAction] = useState<'preview' | 'download' | null>(null)
  const [activeAuthorImageAction, setActiveAuthorImageAction] = useState<
    'preview' | 'download' | null
  >(null)
  const [activeCoverAction, setActiveCoverAction] = useState<'preview' | 'download' | null>(null)

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
    if (!isEditMode || !id) {
      return
    }

    const loadBook = async () => {
      try {
        setIsLoading(true)
        const book = await bookshelfApi.getBook(id)
        setCurrentBook(book)
      } catch (loadError) {
        toast.error(loadError instanceof Error ? loadError.message : 'Could not load this book.')
      } finally {
        setIsLoading(false)
      }
    }

    void loadBook()
  }, [id, isEditMode])

  useEffect(() => {
    if (isEditMode && !currentBook) {
      return
    }

    setForm(currentBook ? bookToFormState(currentBook) : emptyFormState())
    setSelectedBookFile(null)
    setSelectedAuthorImageFile(null)
    setSelectedCoverFile(null)
    setRemoveAuthorImage(false)
    setRemoveCoverImage(false)
  }, [currentBook, isEditMode])

  useEffect(() => {
    let cancelled = false
    let nextPreviewUrl: string | null = null

    async function loadPreview() {
      if (selectedBookFile) {
        nextPreviewUrl = URL.createObjectURL(selectedBookFile)
      } else if (currentBook && canPreviewBook(currentBook.bookMimeType, currentBook.bookFileName)) {
        try {
          const blob = await bookshelfApi.getBookContent(currentBook.id)
          nextPreviewUrl = URL.createObjectURL(blob)
        } catch {
          nextPreviewUrl = null
        }
      }

      if (cancelled) {
        if (nextPreviewUrl) {
          URL.revokeObjectURL(nextPreviewUrl)
        }
        return
      }

      if (bookPreviewUrlRef.current) {
        URL.revokeObjectURL(bookPreviewUrlRef.current)
      }
      bookPreviewUrlRef.current = nextPreviewUrl
      setBookPreviewUrl(nextPreviewUrl)
    }

    void loadPreview()

    return () => {
      cancelled = true
    }
  }, [currentBook, selectedBookFile])

  useEffect(() => {
    let cancelled = false
    let nextPreviewUrl: string | null = null

    async function loadPreview() {
      if (selectedAuthorImageFile) {
        nextPreviewUrl = URL.createObjectURL(selectedAuthorImageFile)
      } else if (currentBook && currentBook.hasAuthorImage && !removeAuthorImage) {
        try {
          const blob = await bookshelfApi.getAuthorImageContent(currentBook.id)
          nextPreviewUrl = URL.createObjectURL(blob)
        } catch {
          nextPreviewUrl = null
        }
      }

      if (cancelled) {
        if (nextPreviewUrl) {
          URL.revokeObjectURL(nextPreviewUrl)
        }
        return
      }

      if (authorImagePreviewUrlRef.current) {
        URL.revokeObjectURL(authorImagePreviewUrlRef.current)
      }
      authorImagePreviewUrlRef.current = nextPreviewUrl
      setAuthorImagePreviewUrl(nextPreviewUrl)
    }

    void loadPreview()

    return () => {
      cancelled = true
    }
  }, [currentBook, removeAuthorImage, selectedAuthorImageFile])

  useEffect(() => {
    let cancelled = false
    let nextPreviewUrl: string | null = null

    async function loadPreview() {
      if (selectedCoverFile) {
        nextPreviewUrl = URL.createObjectURL(selectedCoverFile)
      } else if (currentBook && currentBook.hasCoverImage && !removeCoverImage) {
        try {
          const blob = await bookshelfApi.getCoverImageContent(currentBook.id)
          nextPreviewUrl = URL.createObjectURL(blob)
        } catch {
          nextPreviewUrl = null
        }
      }

      if (cancelled) {
        if (nextPreviewUrl) {
          URL.revokeObjectURL(nextPreviewUrl)
        }
        return
      }

      if (coverPreviewUrlRef.current) {
        URL.revokeObjectURL(coverPreviewUrlRef.current)
      }
      coverPreviewUrlRef.current = nextPreviewUrl
      setCoverPreviewUrl(nextPreviewUrl)
    }

    void loadPreview()

    return () => {
      cancelled = true
    }
  }, [currentBook, removeCoverImage, selectedCoverFile])

  useEffect(() => {
    return () => {
      if (bookPreviewUrlRef.current) {
        URL.revokeObjectURL(bookPreviewUrlRef.current)
      }
      if (authorImagePreviewUrlRef.current) {
        URL.revokeObjectURL(authorImagePreviewUrlRef.current)
      }
      if (coverPreviewUrlRef.current) {
        URL.revokeObjectURL(coverPreviewUrlRef.current)
      }
    }
  }, [])

  function updateField<K extends keyof BookshelfFormState>(key: K, value: BookshelfFormState[K]) {
    setForm((current) => ({ ...current, [key]: value }))

    if (errors[key]) {
      setErrors((current) => {
        const next = { ...current }
        delete next[key]
        return next
      })
    }
  }

  function validate() {
    const nextErrors: BookshelfFormErrors = {}
    const hasExistingBook = Boolean(currentBook?.bookContentUrl)
    const trimmedBookLink = form.bookLink.trim()

    if (!form.author.trim()) {
      nextErrors.author = 'Author is required.'
    }
    if (!form.title.trim()) {
      nextErrors.title = 'Book title is required.'
    }
    if (!form.description.trim()) {
      nextErrors.description = 'Description is required.'
    }
    if (trimmedBookLink && !isValidExternalUrl(trimmedBookLink)) {
      nextErrors.bookLink = 'Purchase link must be a valid http or https URL.'
    }

    if (selectedBookFile) {
      const bookFileError = getBookshelfBookValidationMessage(selectedBookFile)
      if (bookFileError) {
        nextErrors.bookFile = bookFileError
      }
    }
    if (!selectedBookFile && !hasExistingBook) {
      nextErrors.bookFile = 'Please upload the book file before saving.'
    }

    if (selectedAuthorImageFile) {
      const authorImageError = getBookshelfAuthorImageValidationMessage(selectedAuthorImageFile)
      if (authorImageError) {
        nextErrors.authorImage = authorImageError
      }
    }

    if (selectedCoverFile) {
      const coverFileError = getBookshelfCoverValidationMessage(selectedCoverFile)
      if (coverFileError) {
        nextErrors.coverImage = coverFileError
      }
    }

    return nextErrors
  }

  async function handleSave() {
    const validationErrors = validate()
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error('Please correct the highlighted fields.')
      return
    }

    setErrors({})
    setIsSubmitting(true)
    try {
      const input = {
        author: form.author.trim(),
        title: form.title.trim(),
        bookLink: form.bookLink.trim(),
        authorBio: form.authorBio.trim(),
        bookTeaser: form.bookTeaser.trim(),
        description: form.description.trim(),
        removeAuthorImage,
        removeCoverImage,
      }

      if (isEditMode && currentBook) {
        const updated = await update(
          currentBook.id,
          input,
          selectedBookFile ?? undefined,
          selectedAuthorImageFile ?? undefined,
          selectedCoverFile ?? undefined,
        )
        setCurrentBook(updated)
        setSelectedBookFile(null)
        setSelectedAuthorImageFile(null)
        setSelectedCoverFile(null)
        setRemoveAuthorImage(false)
        setRemoveCoverImage(false)
        toast.success('Book updated.')
      } else {
        const created = await create(
          input,
          selectedBookFile ?? undefined,
          selectedAuthorImageFile ?? undefined,
          selectedCoverFile ?? undefined,
        )
        toast.success('Book created.')
        navigate(`/bookshelf/${created.id}/edit`, { replace: true })
      }
    } catch (saveError) {
      toast.error(saveError instanceof Error ? saveError.message : 'Could not save this book.')
    } finally {
      setIsSubmitting(false)
    }
  }

  async function handleDelete() {
    if (!currentBook) {
      return
    }
    setIsDeleting(true)
    try {
      await remove(currentBook.id)
      toast.success('Book deleted.')
      navigate('/bookshelf', { replace: true })
    } catch (deleteError) {
      toast.error(deleteError instanceof Error ? deleteError.message : 'Could not delete this book.')
    } finally {
      setIsDeleting(false)
    }
  }

  function handleBookFileSelect(files: File[]) {
    const [file] = files
    if (!file) {
      return
    }

    const validationMessage = getBookshelfBookValidationMessage(file)
    if (validationMessage) {
      setErrors((current) => ({ ...current, bookFile: validationMessage }))
      toast.error(validationMessage)
      return
    }

    setSelectedBookFile(file)
    if (errors.bookFile) {
      setErrors((current) => {
        const next = { ...current }
        delete next.bookFile
        return next
      })
    }
  }

  function handleAuthorImageFileSelect(files: File[]) {
    const [file] = files
    if (!file) {
      return
    }

    const validationMessage = getBookshelfAuthorImageValidationMessage(file)
    if (validationMessage) {
      setErrors((current) => ({ ...current, authorImage: validationMessage }))
      toast.error(validationMessage)
      return
    }

    setSelectedAuthorImageFile(file)
    setRemoveAuthorImage(false)
    if (errors.authorImage) {
      setErrors((current) => {
        const next = { ...current }
        delete next.authorImage
        return next
      })
    }
  }

  function handleCoverFileSelect(files: File[]) {
    const [file] = files
    if (!file) {
      return
    }

    const validationMessage = getBookshelfCoverValidationMessage(file)
    if (validationMessage) {
      setErrors((current) => ({ ...current, coverImage: validationMessage }))
      toast.error(validationMessage)
      return
    }

    setSelectedCoverFile(file)
    setRemoveCoverImage(false)
    if (errors.coverImage) {
      setErrors((current) => {
        const next = { ...current }
        delete next.coverImage
        return next
      })
    }
  }

  function handleRemoveSelectedBook() {
    setSelectedBookFile(null)
  }

  function handleRemoveAuthorImage() {
    if (selectedAuthorImageFile) {
      setSelectedAuthorImageFile(null)
      return
    }
    if (currentBook?.hasAuthorImage) {
      setRemoveAuthorImage(true)
    }
  }

  function handleRestoreAuthorImage() {
    setRemoveAuthorImage(false)
  }

  function handleRemoveCover() {
    if (selectedCoverFile) {
      setSelectedCoverFile(null)
      return
    }
    if (currentBook?.hasCoverImage) {
      setRemoveCoverImage(true)
    }
  }

  function handleRestoreCover() {
    setRemoveCoverImage(false)
  }

  async function handlePreviewBook() {
    const fileName = selectedBookFile?.name || currentBook?.bookFileName
    const mimeType = selectedBookFile?.type || currentBook?.bookMimeType || ''

    if (!fileName || !canPreviewBook(mimeType, fileName)) {
      return
    }

    setActiveBookAction('preview')
    try {
      const previewUrl = await createTemporaryBookshelfObjectUrl(
        'book',
        currentBook?.id,
        selectedBookFile,
      )
      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!selectedBookFile) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error('Could not preview this file.')
    } finally {
      setActiveBookAction(null)
    }
  }

  async function handleDownloadBook() {
    setActiveBookAction('download')
    try {
      const downloadUrl = await createTemporaryBookshelfObjectUrl(
        'book',
        currentBook?.id,
        selectedBookFile,
      )
      const fileName = selectedBookFile?.name || currentBook?.bookFileName || 'book-file'
      triggerFileDownload(downloadUrl, fileName)
      if (!selectedBookFile) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error('Could not download this file.')
    } finally {
      setActiveBookAction(null)
    }
  }

  async function handlePreviewAuthorImage() {
    if (!selectedAuthorImageFile && (!currentBook?.hasAuthorImage || removeAuthorImage)) {
      return
    }

    setActiveAuthorImageAction('preview')
    try {
      const previewUrl = await createTemporaryBookshelfObjectUrl(
        'authorImage',
        currentBook?.id,
        selectedAuthorImageFile,
      )
      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!selectedAuthorImageFile) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error('Could not preview this author image.')
    } finally {
      setActiveAuthorImageAction(null)
    }
  }

  async function handleDownloadAuthorImage() {
    if (!selectedAuthorImageFile && (!currentBook?.hasAuthorImage || removeAuthorImage)) {
      return
    }

    setActiveAuthorImageAction('download')
    try {
      const downloadUrl = await createTemporaryBookshelfObjectUrl(
        'authorImage',
        currentBook?.id,
        selectedAuthorImageFile,
      )
      const fileName =
        selectedAuthorImageFile?.name || currentBook?.authorImageFileName || 'author-image'
      triggerFileDownload(downloadUrl, fileName)
      if (!selectedAuthorImageFile) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error('Could not download this author image.')
    } finally {
      setActiveAuthorImageAction(null)
    }
  }

  async function handlePreviewCover() {
    if (!selectedCoverFile && (!currentBook?.hasCoverImage || removeCoverImage)) {
      return
    }

    setActiveCoverAction('preview')
    try {
      const previewUrl = await createTemporaryBookshelfObjectUrl(
        'cover',
        currentBook?.id,
        selectedCoverFile,
      )
      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!selectedCoverFile) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error('Could not preview this cover image.')
    } finally {
      setActiveCoverAction(null)
    }
  }

  async function handleDownloadCover() {
    if (!selectedCoverFile && (!currentBook?.hasCoverImage || removeCoverImage)) {
      return
    }

    setActiveCoverAction('download')
    try {
      const downloadUrl = await createTemporaryBookshelfObjectUrl(
        'cover',
        currentBook?.id,
        selectedCoverFile,
      )
      const fileName = selectedCoverFile?.name || currentBook?.coverImageFileName || 'cover-image'
      triggerFileDownload(downloadUrl, fileName)
      if (!selectedCoverFile) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error('Could not download this cover image.')
    } finally {
      setActiveCoverAction(null)
    }
  }

  function handleOpenPurchaseLink() {
    const bookLink = form.bookLink.trim()
    if (!isValidExternalUrl(bookLink)) {
      return
    }
    window.open(bookLink, '_blank', 'noopener,noreferrer')
  }

  if (isLoading) {
    return (
      <CmsAppShell activeKey="bookshelf">
        <div className={styles.loaderWrap}>
          <Loader label="Loading book..." />
        </div>
      </CmsAppShell>
    )
  }

  const bookName = selectedBookFile?.name || currentBook?.bookFileName || ''
  const bookMimeType = selectedBookFile?.type || currentBook?.bookMimeType || ''
  const bookSize = selectedBookFile?.size ?? currentBook?.bookFileSize ?? 0
  const canPreviewSelectedBook = canPreviewBook(bookMimeType, bookName)
  const shouldShowBookCard = Boolean(selectedBookFile || currentBook?.bookContentUrl)

  const authorImageName =
    selectedAuthorImageFile?.name || currentBook?.authorImageFileName || ''
  const authorImageSize =
    selectedAuthorImageFile?.size ?? currentBook?.authorImageFileSize ?? 0
  const shouldShowAuthorImageCard = Boolean(
    selectedAuthorImageFile || (currentBook?.hasAuthorImage && !removeAuthorImage),
  )

  const coverName = selectedCoverFile?.name || currentBook?.coverImageFileName || ''
  const coverSize = selectedCoverFile?.size ?? currentBook?.coverImageFileSize ?? 0
  const shouldShowCoverCard = Boolean(
    selectedCoverFile || (currentBook?.hasCoverImage && !removeCoverImage),
  )

  const canOpenPurchaseLink = isValidExternalUrl(form.bookLink.trim())
  const errorMessages = Array.from(new Set(Object.values(errors).filter(Boolean)))

  return (
    <CmsAppShell activeKey="bookshelf">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: 'Bookshelf', to: '/bookshelf' },
            { label: isEditMode ? 'Edit Book' : 'Add Book' },
          ]}
        />

        <div className={styles.header}>
          <div className={styles.headerText}>
            <h1>{isEditMode ? 'Edit Book' : 'Add Book'}</h1>
            <p>
              Manage the uploaded book file, teaser copy, purchase link, author profile,
              and optional artwork from one editor.
            </p>
          </div>
          <button type="button" className={styles.backLink} onClick={() => navigate('/bookshelf')}>
            Back to bookshelf
          </button>
        </div>

        {errorMessages.length > 0 ? (
          <div className={styles.errorSummary} role="alert">
            <p>Please correct the highlighted fields.</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        ) : null}

        <div className={styles.layout}>
          <div className={styles.mainColumn}>
            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>Book Details</h2>
              </div>

              <div className={styles.fieldGrid}>
                <label className={styles.field}>
                  <span>Author</span>
                  <input
                    type="text"
                    value={form.author}
                    placeholder="Enter the book author"
                    onChange={(event) => updateField('author', event.target.value)}
                  />
                  {errors.author ? <p className={styles.fieldError}>{errors.author}</p> : null}
                </label>

                <label className={styles.field}>
                  <span>Book Title</span>
                  <input
                    type="text"
                    value={form.title}
                    placeholder="Enter the book title"
                    onChange={(event) => updateField('title', event.target.value)}
                  />
                  {errors.title ? <p className={styles.fieldError}>{errors.title}</p> : null}
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Book Teaser (Optional)</span>
                  <textarea
                    rows={3}
                    value={form.bookTeaser}
                    placeholder="Add a short preview line for the bookshelf listing"
                    onChange={(event) => updateField('bookTeaser', event.target.value)}
                  />
                  {errors.bookTeaser ? (
                    <p className={styles.fieldError}>{errors.bookTeaser}</p>
                  ) : null}
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Purchase Link</span>
                  <input
                    type="url"
                    value={form.bookLink}
                    placeholder="https://example.com/buy-the-book"
                    onChange={(event) => updateField('bookLink', event.target.value)}
                  />
                  {errors.bookLink ? <p className={styles.fieldError}>{errors.bookLink}</p> : null}
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Description</span>
                  <textarea
                    rows={5}
                    value={form.description}
                    placeholder="Add the full description for this book"
                    onChange={(event) => updateField('description', event.target.value)}
                  />
                  {errors.description ? (
                    <p className={styles.fieldError}>{errors.description}</p>
                  ) : null}
                </label>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>Author Profile</h2>
                <span>Optional image, max {BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB}MB</span>
              </div>

              <div className={styles.fieldGrid}>
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Author Bio (Optional)</span>
                  <textarea
                    rows={5}
                    value={form.authorBio}
                    placeholder="Add a short author biography"
                    onChange={(event) => updateField('authorBio', event.target.value)}
                  />
                  {errors.authorBio ? <p className={styles.fieldError}>{errors.authorBio}</p> : null}
                </label>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  accept={BOOKSHELF_AUTHOR_IMAGE_ACCEPT}
                  icon={<UploadCloudIcon />}
                  label="Drag and drop the author image here"
                  hint={`Supports PNG, JPG, WEBP, and SVG up to ${BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB}MB`}
                  onFiles={handleAuthorImageFileSelect}
                />
                {errors.authorImage ? (
                  <p className={styles.fieldError}>{errors.authorImage}</p>
                ) : null}

                {removeAuthorImage && currentBook?.hasAuthorImage && !selectedAuthorImageFile ? (
                  <div className={styles.metaCard}>
                    <span className={styles.metaLabel}>Pending change</span>
                    <strong>The current author image will be removed when you save.</strong>
                    <button
                      type="button"
                      className={styles.actionButton}
                      onClick={handleRestoreAuthorImage}
                    >
                      Keep current author image
                    </button>
                  </div>
                ) : null}

                {shouldShowAuthorImageCard ? (
                  <article className={styles.documentCard}>
                    <div className={styles.documentPreviewCard} aria-hidden="true">
                      {authorImagePreviewUrl ? (
                        <img
                          src={authorImagePreviewUrl}
                          alt={authorImageName || 'Author image preview'}
                          className={styles.documentPreviewFrame}
                        />
                      ) : (
                        <>
                          <ImageIcon />
                          <span className={styles.documentPreviewBadge}>
                            {resolveImageTypeLabel(authorImageName)}
                          </span>
                        </>
                      )}
                    </div>

                    <div className={styles.documentContent}>
                      <h3>{authorImageName || 'Author image'}</h3>
                      <p className={styles.documentMeta}>
                        {buildDocumentMeta(
                          Boolean(selectedAuthorImageFile),
                          resolveImageTypeLabel(authorImageName),
                          authorImageSize,
                        )}
                      </p>

                      <div className={styles.documentActions}>
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeAuthorImageAction !== null}
                          onClick={() => void handlePreviewAuthorImage()}
                        >
                          {activeAuthorImageAction === 'preview' ? 'Loading...' : 'Preview'}
                        </button>
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeAuthorImageAction !== null}
                          onClick={() => void handleDownloadAuthorImage()}
                        >
                          {activeAuthorImageAction === 'download' ? 'Loading...' : 'Download'}
                        </button>
                        <button
                          type="button"
                          className={styles.dangerButton}
                          onClick={handleRemoveAuthorImage}
                        >
                          {selectedAuthorImageFile
                            ? 'Remove selected image'
                            : 'Remove current author image'}
                        </button>
                      </div>
                    </div>
                  </article>
                ) : null}
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>Book Upload</h2>
                <span>Max {BOOKSHELF_BOOK_MAX_FILE_SIZE_MB}MB / file</span>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  accept={BOOKSHELF_BOOK_ACCEPT}
                  icon={<UploadCloudIcon />}
                  label="Drag and drop the book file here"
                  hint={`Supports PDF, EPUB, DOC, and DOCX up to ${BOOKSHELF_BOOK_MAX_FILE_SIZE_MB}MB`}
                  onFiles={handleBookFileSelect}
                />
                {errors.bookFile ? <p className={styles.fieldError}>{errors.bookFile}</p> : null}

                {shouldShowBookCard ? (
                  <article className={styles.documentCard}>
                    <div className={styles.documentPreviewCard} aria-hidden="true">
                      {canPreviewSelectedBook && bookPreviewUrl ? (
                        <iframe
                          src={bookPreviewUrl}
                          title={bookName}
                          className={styles.documentPreviewFrame}
                        />
                      ) : (
                        <>
                          <DocumentIcon />
                          <span className={styles.documentPreviewBadge}>
                            {resolveDocumentTypeLabel(bookMimeType, bookName)}
                          </span>
                        </>
                      )}
                    </div>

                    <div className={styles.documentContent}>
                      <h3>{form.title || bookName || 'Uploaded book'}</h3>
                      <p className={styles.documentMeta}>
                        {buildDocumentMeta(
                          Boolean(selectedBookFile),
                          resolveDocumentTypeLabel(bookMimeType, bookName),
                          bookSize,
                        )}
                      </p>

                      <div className={styles.documentActions}>
                        {canPreviewSelectedBook ? (
                          <button
                            type="button"
                            className={styles.actionButton}
                            disabled={activeBookAction !== null}
                            onClick={() => void handlePreviewBook()}
                          >
                            {activeBookAction === 'preview' ? 'Loading...' : 'Preview'}
                          </button>
                        ) : null}
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeBookAction !== null}
                          onClick={() => void handleDownloadBook()}
                        >
                          {activeBookAction === 'download' ? 'Loading...' : 'Download'}
                        </button>
                        {selectedBookFile ? (
                          <button
                            type="button"
                            className={styles.dangerButton}
                            onClick={handleRemoveSelectedBook}
                          >
                            Remove selected file
                          </button>
                        ) : null}
                      </div>
                    </div>
                  </article>
                ) : null}
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>Cover Image</h2>
                <span>Optional, max {BOOKSHELF_COVER_MAX_FILE_SIZE_MB}MB</span>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  accept={BOOKSHELF_COVER_ACCEPT}
                  icon={<UploadCloudIcon />}
                  label="Drag and drop the cover image here"
                  hint={`Supports PNG, JPG, WEBP, and SVG up to ${BOOKSHELF_COVER_MAX_FILE_SIZE_MB}MB`}
                  onFiles={handleCoverFileSelect}
                />
                {errors.coverImage ? <p className={styles.fieldError}>{errors.coverImage}</p> : null}

                {removeCoverImage && currentBook?.hasCoverImage && !selectedCoverFile ? (
                  <div className={styles.metaCard}>
                    <span className={styles.metaLabel}>Pending change</span>
                    <strong>The current cover image will be removed when you save.</strong>
                    <button type="button" className={styles.actionButton} onClick={handleRestoreCover}>
                      Keep current cover
                    </button>
                  </div>
                ) : null}

                {shouldShowCoverCard ? (
                  <article className={styles.documentCard}>
                    <div className={styles.documentPreviewCard} aria-hidden="true">
                      {coverPreviewUrl ? (
                        <img
                          src={coverPreviewUrl}
                          alt={coverName || 'Cover preview'}
                          className={styles.documentPreviewFrame}
                        />
                      ) : (
                        <>
                          <ImageIcon />
                          <span className={styles.documentPreviewBadge}>
                            {resolveImageTypeLabel(coverName)}
                          </span>
                        </>
                      )}
                    </div>

                    <div className={styles.documentContent}>
                      <h3>{coverName || 'Cover image'}</h3>
                      <p className={styles.documentMeta}>
                        {buildDocumentMeta(
                          Boolean(selectedCoverFile),
                          resolveImageTypeLabel(coverName),
                          coverSize,
                        )}
                      </p>

                      <div className={styles.documentActions}>
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeCoverAction !== null}
                          onClick={() => void handlePreviewCover()}
                        >
                          {activeCoverAction === 'preview' ? 'Loading...' : 'Preview'}
                        </button>
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeCoverAction !== null}
                          onClick={() => void handleDownloadCover()}
                        >
                          {activeCoverAction === 'download' ? 'Loading...' : 'Download'}
                        </button>
                        <button type="button" className={styles.dangerButton} onClick={handleRemoveCover}>
                          {selectedCoverFile ? 'Remove selected image' : 'Remove current cover'}
                        </button>
                      </div>
                    </div>
                  </article>
                ) : null}
              </div>
            </section>

            <div className={styles.actionBar}>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={isSubmitting}
                onClick={() => void handleSave()}
              >
                {isSubmitting ? 'Loading...' : isEditMode ? 'Save Changes' : 'Create Book'}
              </button>
              {isEditMode ? (
                <button
                  type="button"
                  className={styles.deleteButton}
                  disabled={isDeleting || isSubmitting}
                  onClick={() => void handleDelete()}
                >
                  {isDeleting ? 'Loading...' : 'Delete Book'}
                </button>
              ) : null}
            </div>
          </div>

          <aside className={styles.sideColumn}>
            <div className={styles.metaCard}>
              <span className={styles.metaLabel}>Book File</span>
              <strong>
                {selectedBookFile
                  ? 'New file selected'
                  : currentBook?.bookFileName || 'No file uploaded yet'}
              </strong>
              <span className={styles.metaLabel}>Author Image</span>
              <strong>
                {selectedAuthorImageFile
                  ? 'New author image selected'
                  : removeAuthorImage
                    ? 'Will be removed'
                    : currentBook?.hasAuthorImage
                      ? 'Saved author image'
                      : 'No author image'}
              </strong>
              <span className={styles.metaLabel}>Cover Image</span>
              <strong>
                {selectedCoverFile
                  ? 'New cover selected'
                  : removeCoverImage
                    ? 'Will be removed'
                    : currentBook?.hasCoverImage
                      ? 'Saved cover image'
                      : 'No cover image'}
              </strong>
              <span className={styles.metaLabel}>Purchase Link</span>
              <strong>{form.bookLink.trim() ? 'Link added' : 'No purchase link'}</strong>
              {canOpenPurchaseLink ? (
                <button
                  type="button"
                  className={styles.actionButton}
                  onClick={handleOpenPurchaseLink}
                >
                  Open purchase link
                </button>
              ) : null}
            </div>

            {currentBook ? (
              <div className={styles.metaCard}>
                <span className={styles.metaLabel}>Created</span>
                <strong>{dateFormatter.format(new Date(currentBook.createdAt))}</strong>
                <span className={styles.metaLabel}>Last updated</span>
                <strong>{dateFormatter.format(new Date(currentBook.updatedAt))}</strong>
              </div>
            ) : null}

            <div className={styles.helpCard}>
              <h3>Quick Tip</h3>
              <p>
                Use clear teaser text and a concise author bio so staff can scan the bookshelf
                quickly and choose the right title with confidence.
              </p>
            </div>
          </aside>
        </div>
      </div>
    </CmsAppShell>
  )
}

async function createTemporaryBookshelfObjectUrl(
  kind: BookshelfPreviewKind,
  bookId?: string,
  file?: File | null,
) {
  if (file) {
    return URL.createObjectURL(file)
  }
  if (!bookId) {
    throw new Error('File URL unavailable')
  }

  switch (kind) {
    case 'book': {
      const blob = await bookshelfApi.getBookContent(bookId)
      return URL.createObjectURL(blob)
    }
    case 'authorImage': {
      const blob = await bookshelfApi.getAuthorImageContent(bookId)
      return URL.createObjectURL(blob)
    }
    case 'cover': {
      const blob = await bookshelfApi.getCoverImageContent(bookId)
      return URL.createObjectURL(blob)
    }
    default:
      throw new Error('Unsupported preview type')
  }
}

function buildDocumentMeta(isPendingUpload: boolean, fileTypeLabel: string, fileSize?: number) {
  const parts = [
    isPendingUpload ? 'Pending upload' : 'Saved file',
    fileTypeLabel,
    typeof fileSize === 'number' && fileSize > 0 ? formatFileSize(fileSize) : '',
  ].filter(Boolean)

  return parts.join(' | ')
}

function getBookshelfBookValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateBookshelfBookFile(file)
  if (!validationError) {
    return null
  }

  if (validationError === 'file-too-large') {
    return `This book file exceeds the ${BOOKSHELF_BOOK_MAX_FILE_SIZE_MB}MB limit.`
  }

  return 'Only PDF, EPUB, DOC, and DOCX files are supported for books.'
}

function getBookshelfAuthorImageValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateBookshelfAuthorImageFile(file)
  if (!validationError) {
    return null
  }

  if (validationError === 'file-too-large') {
    return `This author image exceeds the ${BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB}MB limit.`
  }

  return 'Only PNG, JPG, WEBP, and SVG files are supported for author images.'
}

function getBookshelfCoverValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateBookshelfCoverFile(file)
  if (!validationError) {
    return null
  }

  if (validationError === 'file-too-large') {
    return `This cover image exceeds the ${BOOKSHELF_COVER_MAX_FILE_SIZE_MB}MB limit.`
  }

  return 'Only PNG, JPG, WEBP, and SVG files are supported for cover images.'
}

function isValidExternalUrl(value: string) {
  if (!value.trim()) {
    return false
  }

  try {
    const parsed = new URL(value)
    return parsed.protocol === 'http:' || parsed.protocol === 'https:'
  } catch {
    return false
  }
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

function canPreviewBook(mimeType: string, fileName: string) {
  const normalizedMimeType = mimeType.trim().toLowerCase()
  const extension = fileName.split('.').pop()?.trim().toLowerCase() ?? ''

  return (
    normalizedMimeType.includes('pdf') ||
    normalizedMimeType.startsWith('image/') ||
    ['pdf', 'png', 'jpg', 'jpeg', 'webp', 'svg'].includes(extension)
  )
}

function resolveDocumentTypeLabel(mimeType: string, fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (extension) {
    return extension.length > 5 ? extension.slice(0, 5) : extension
  }

  const normalizedMimeType = mimeType.trim().toLowerCase()
  if (normalizedMimeType.includes('pdf')) {
    return 'PDF'
  }
  if (normalizedMimeType.includes('epub')) {
    return 'EPUB'
  }
  if (normalizedMimeType.includes('word') || normalizedMimeType.includes('document')) {
    return 'DOC'
  }

  return 'FILE'
}

function resolveImageTypeLabel(fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (!extension) {
    return 'IMG'
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

function UploadCloudIcon() {
  return (
    <svg viewBox="0 0 24 24" width="22" height="22" fill="none" aria-hidden="true">
      <path
        d="M8.5 18H7a4 4 0 1 1 .7-7.94A5.5 5.5 0 0 1 18.38 12H19a3 3 0 1 1 0 6h-2.5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M12 10v8m0-8 3 3m-3-3-3 3"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function DocumentIcon() {
  return (
    <svg viewBox="0 0 24 24" width="28" height="28" fill="none" aria-hidden="true">
      <path
        d="M8 3.5h6l4 4V20.5H8z"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
      <path d="M14 3.5v4h4" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
    </svg>
  )
}

function ImageIcon() {
  return (
    <svg viewBox="0 0 24 24" width="28" height="28" fill="none" aria-hidden="true">
      <rect x="4" y="5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="1.5" />
      <path
        d="M7 16l3.2-3.2 2.3 2.3 2.8-3.1L18 16"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <circle cx="9" cy="9" r="1.2" fill="currentColor" />
    </svg>
  )
}
