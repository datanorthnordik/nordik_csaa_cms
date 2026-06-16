import { useEffect, useMemo, useState, type ChangeEvent } from 'react'
import toast from 'react-hot-toast'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import {
  booksApi,
  type BookDetail,
  type BookSummary,
  type BookFieldInputType,
  type BookFieldPlacement,
  type BookSubmission,
  type BookSubmissionSaveInput,
  type BookVersionDetail,
  type BookVersionSaveInput,
} from '../api/booksApi'
import { generateBookPdf } from '../lib/bookPdf'
import styles from '../styles/BooksPage.module.css'

const DEFAULT_LAYOUT_SETTINGS = {
  content_mask: {
    x: 54,
    y: 92,
    width: 392,
    height: 484,
    background_color: '#ffffff',
  },
  heading_area: {
    x: 78,
    y: 114,
    width: 316,
    height: 86,
    font_size: 19,
    line_height: 1.2,
    text_align: 'left',
  },
  body_area: {
    x: 78,
    y: 214,
    width: 280,
    height: 314,
    font_size: 11,
    line_height: 1.35,
    text_align: 'left',
  },
  image_area: {
    x: 316,
    y: 422,
    width: 108,
    height: 108,
  },
  section_mask: {
    x: 70,
    y: 228,
    width: 360,
    height: 114,
    background_color: '#ffffff',
  },
  section_title_area: {
    x: 90,
    y: 248,
    width: 320,
    height: 74,
    font_size: 28,
    line_height: 1.1,
    text_align: 'center',
  },
}

type BookFormState = {
  title: string
  description: string
  adminEmails: string
}

type EditableSection = {
  id?: number
  name: string
  sourceStartPage: string
  sourceEndPage: string
}

type EditableField = {
  id?: number
  label: string
  inputType: BookFieldInputType
  placement: BookFieldPlacement
  showLabel: boolean
  isRequired: boolean
  isEmailField: boolean
}

type VersionFormState = {
  id?: number
  sourcePageCount: string
  contentTemplatePageNumber: string
  sectionTemplatePageNumber: string
  allowPageImage: boolean
  allowNewSections: boolean
  activateImmediately: boolean
  layoutSettingsText: string
  sections: EditableSection[]
  fields: EditableField[]
  sourcePdfFile: File | null
}

type SubmissionCardProps = {
  submission: BookSubmission
  versionDetail: BookVersionDetail
  busy: boolean
  onSave: (submissionId: number, input: BookSubmissionSaveInput, imageFile?: File | null) => Promise<void>
  onApprove: (submissionId: number) => Promise<void>
  onReject: (submissionId: number, rejectionReason: string) => Promise<void>
}

export function BooksPage() {
  const [books, setBooks] = useState<BookSummary[]>([])
  const [selectedBookId, setSelectedBookId] = useState<number | null>(null)
  const [selectedBook, setSelectedBook] = useState<BookDetail | null>(null)
  const [selectedVersionId, setSelectedVersionId] = useState<number | null>(null)
  const [selectedVersion, setSelectedVersion] = useState<BookVersionDetail | null>(null)
  const [submissions, setSubmissions] = useState<BookSubmission[]>([])
  const [submissionStatusFilter, setSubmissionStatusFilter] = useState('')
  const [bookForm, setBookForm] = useState<BookFormState>({
    title: '',
    description: '',
    adminEmails: '',
  })
  const [versionForm, setVersionForm] = useState<VersionFormState>(buildEmptyVersionForm())
  const [isCreatingNewVersion, setIsCreatingNewVersion] = useState(false)
  const [isBooksLoading, setIsBooksLoading] = useState(true)
  const [isBookSaving, setIsBookSaving] = useState(false)
  const [isVersionSaving, setIsVersionSaving] = useState(false)
  const [isVersionLoading, setIsVersionLoading] = useState(false)
  const [isSubmissionsLoading, setIsSubmissionsLoading] = useState(false)
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false)
  const [busySubmissionId, setBusySubmissionId] = useState<number | null>(null)

  const selectedVersionSummary = useMemo(
    () => selectedBook?.versions.find((version) => version.id === selectedVersionId) ?? null,
    [selectedBook, selectedVersionId],
  )

  useEffect(() => {
    void loadBooks()
  }, [])

  useEffect(() => {
    if (!selectedBookId) {
      return
    }
    void loadBookDetail(selectedBookId, { preserveVersionSelection: true })
  }, [selectedBookId])

  useEffect(() => {
    if (!selectedBookId || !selectedVersionId || !selectedBook) {
      return
    }
    void loadVersionDetail(selectedBookId, selectedVersionId)
  }, [selectedBookId, selectedVersionId, submissionStatusFilter])

  async function loadBooks() {
    try {
      setIsBooksLoading(true)
      const bookSummaries = await booksApi.listBooks()
      setBooks(bookSummaries)

      const nextSelectedBookId = selectedBookId ?? bookSummaries[0]?.id ?? null
      if (nextSelectedBookId) {
        setSelectedBookId(nextSelectedBookId)
      } else {
        setSelectedBook(null)
        setSelectedVersion(null)
        setSelectedVersionId(null)
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load books.')
    } finally {
      setIsBooksLoading(false)
    }
  }

  async function loadBookDetail(
    bookId: number,
    options: { preserveVersionSelection?: boolean } = {},
  ) {
    try {
      const detail = await booksApi.getBook(bookId)
      setSelectedBook(detail)
      setBookForm({
        title: detail.title,
        description: detail.description,
        adminEmails: detail.adminNotificationEmails.join(', '),
      })

      const nextVersionId =
        options.preserveVersionSelection && selectedVersionId && detail.versions.some((item) => item.id === selectedVersionId)
          ? selectedVersionId
          : detail.activeVersionId ?? detail.versions[0]?.id ?? null

      setSelectedVersionId(nextVersionId)
      if (!nextVersionId) {
        setSelectedVersion(null)
        setVersionForm(buildEmptyVersionForm())
        setSubmissions([])
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load the selected book.')
    }
  }

  async function loadVersionDetail(bookId: number, versionId: number) {
    try {
      setIsVersionLoading(true)
      setIsSubmissionsLoading(true)
      const detail = await booksApi.getVersion(bookId, versionId)
      setSelectedVersion(detail)
      setSubmissions(
        submissionStatusFilter
          ? await booksApi.listSubmissions(bookId, versionId, submissionStatusFilter)
          : await booksApi.listSubmissions(bookId, versionId),
      )
      if (!isCreatingNewVersion) {
        setVersionForm(versionDetailToForm(detail))
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load the selected version.')
    } finally {
      setIsVersionLoading(false)
      setIsSubmissionsLoading(false)
    }
  }

  async function handleSaveBook() {
    try {
      setIsBookSaving(true)
      const payload = {
        title: bookForm.title.trim(),
        description: bookForm.description.trim(),
        adminNotificationEmails: bookForm.adminEmails
          .split(',')
          .map((item) => item.trim())
          .filter(Boolean),
      }

      if (!payload.title) {
        toast.error('Book title is required.')
        return
      }

      let bookId = selectedBook?.id ?? null
      if (bookId) {
        await booksApi.updateBook(bookId, payload)
      } else {
        const created = await booksApi.createBook(payload)
        bookId = created.id
        setSelectedBookId(bookId)
      }

      await loadBooks()
      await loadBookDetail(bookId)
      toast.success('Book details saved.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to save book details.')
    } finally {
      setIsBookSaving(false)
    }
  }

  async function handleSaveVersion() {
    if (!selectedBook?.id) {
      toast.error('Save the book first before configuring versions.')
      return
    }

    try {
      setIsVersionSaving(true)
      const payload = buildVersionPayload(versionForm)
      if (!payload) {
        return
      }

      if (isCreatingNewVersion || !versionForm.id) {
        const sourcePdfFile =
          versionForm.sourcePdfFile ?? (await buildSourcePdfCloneFile(selectedBook.id, selectedVersion))
        if (!sourcePdfFile) {
          toast.error('A source PDF is required for a new version.')
          return
        }

        const created = await booksApi.createVersion(selectedBook.id, payload, sourcePdfFile)
        setIsCreatingNewVersion(false)
        await loadBooks()
        await loadBookDetail(selectedBook.id)
        setSelectedVersionId(created.id)
        toast.success(`Version ${created.version_number} created.`)
      } else {
        await booksApi.updateVersion(selectedBook.id, versionForm.id, payload, versionForm.sourcePdfFile)
        await loadBooks()
        await loadBookDetail(selectedBook.id)
        await loadVersionDetail(selectedBook.id, versionForm.id)
        toast.success('Version updated.')
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to save the version.')
    } finally {
      setIsVersionSaving(false)
    }
  }

  async function handleActivateVersion(versionId: number) {
    if (!selectedBook?.id) {
      return
    }
    try {
      await booksApi.activateVersion(selectedBook.id, versionId)
      await loadBooks()
      await loadBookDetail(selectedBook.id)
      await loadVersionDetail(selectedBook.id, versionId)
      toast.success('Version activated.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to activate the version.')
    }
  }

  async function handleGeneratePdf(versionOverride?: BookVersionDetail) {
    if (!selectedBook?.id || !selectedVersionId) {
      return
    }

    try {
      setIsGeneratingPdf(true)
      const detail = versionOverride ?? (await booksApi.getVersion(selectedBook.id, selectedVersionId))
      const sourceBlob = await booksApi.fetchSourcePdfBlob(selectedBook.id, detail.id)
      const sourceBytes = new Uint8Array(await sourceBlob.arrayBuffer())
      const generatedBlob = await generateBookPdf({
        version: detail,
        sourcePdfBytes: sourceBytes,
        fetchImageBytes: async (submission) => {
          if (!submission.image?.fetchUrl) {
            return null
          }
          const blob = await booksApi.fetchSubmissionImage(submission.image.fetchUrl)
          return new Uint8Array(await blob.arrayBuffer())
        },
      })
      const fileName = `${slugify(selectedBook.title)}-version-${detail.versionNumber}.pdf`
      await booksApi.uploadGeneratedPdf(selectedBook.id, detail.id, generatedBlob, fileName)
      await loadBooks()
      await loadBookDetail(selectedBook.id)
      await loadVersionDetail(selectedBook.id, detail.id)
      toast.success('Generated PDF uploaded.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to generate the PDF.')
    } finally {
      setIsGeneratingPdf(false)
    }
  }

  function handleCreateNewBook() {
    setSelectedBookId(null)
    setSelectedBook(null)
    setSelectedVersionId(null)
    setSelectedVersion(null)
    setSubmissions([])
    setBookForm({
      title: '',
      description: '',
      adminEmails: '',
    })
    setVersionForm(buildEmptyVersionForm())
    setIsCreatingNewVersion(false)
  }

  function handleCreateNewVersion() {
    setIsCreatingNewVersion(true)
    setVersionForm(buildEmptyVersionForm(selectedVersion))
  }

  async function handleSaveSubmission(
    submissionId: number,
    input: BookSubmissionSaveInput,
    imageFile?: File | null,
  ) {
    if (!selectedBook?.id || !selectedVersionId) {
      return
    }
    try {
      setBusySubmissionId(submissionId)
      await booksApi.updateSubmission(selectedBook.id, submissionId, input, imageFile)
      await loadVersionDetail(selectedBook.id, selectedVersionId)
      toast.success('Submission updated.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to update the submission.')
    } finally {
      setBusySubmissionId(null)
    }
  }

  async function handleApproveSubmission(submissionId: number) {
    if (!selectedBook?.id || !selectedVersionId) {
      return
    }
    try {
      setBusySubmissionId(submissionId)
      await booksApi.approveSubmission(selectedBook.id, submissionId)
      const refreshedVersion = await booksApi.getVersion(selectedBook.id, selectedVersionId)
      await handleGeneratePdf(refreshedVersion)
      await loadVersionDetail(selectedBook.id, selectedVersionId)
      toast.success('Submission approved and PDF regenerated.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to approve the submission.')
    } finally {
      setBusySubmissionId(null)
    }
  }

  async function handleRejectSubmission(submissionId: number, rejectionReason: string) {
    if (!selectedBook?.id || !selectedVersionId) {
      return
    }
    try {
      setBusySubmissionId(submissionId)
      await booksApi.rejectSubmission(selectedBook.id, submissionId, rejectionReason)
      await loadVersionDetail(selectedBook.id, selectedVersionId)
      toast.success('Submission rejected.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to reject the submission.')
    } finally {
      setBusySubmissionId(null)
    }
  }

  return (
    <CmsAppShell activeKey="books">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: 'Books' }]} />

        <div className={styles.hero}>
          <div>
            <h1 className={styles.title}>Book Builder</h1>
            <p className={styles.subtitle}>
              Manage book metadata, version schemas, section mappings, survivor submissions,
              and generated PDFs from one CMS workspace.
            </p>
          </div>
          <div className={styles.heroActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => void loadBooks()}>
              Refresh
            </button>
            <button type="button" className={styles.primaryButton} onClick={handleCreateNewBook}>
              New Book
            </button>
          </div>
        </div>

        <div className={styles.layout}>
          <aside className={styles.sidebar}>
            <div className={styles.panelHeader}>
              <h2>Books</h2>
              <span>{books.length}</span>
            </div>

            {isBooksLoading ? (
              <div className={styles.loaderPanel}>
                <Loader />
              </div>
            ) : books.length > 0 ? (
              <div className={styles.bookList}>
                {books
                  .slice()
                  .sort((left, right) => left.title.localeCompare(right.title))
                  .map((book) => (
                    <button
                      key={book.id}
                      type="button"
                      className={`${styles.bookListItem} ${selectedBookId === book.id ? styles.bookListItemActive : ''}`}
                      onClick={() => {
                        setSelectedBookId(book.id)
                        setIsCreatingNewVersion(false)
                      }}
                    >
                      <strong>{book.title}</strong>
                      <span>{book.pendingSubmissionCount} pending</span>
                      <span>
                        {book.activeVersionNumber
                          ? `Active version ${book.activeVersionNumber}`
                          : 'No active version'}
                      </span>
                    </button>
                  ))}
              </div>
            ) : (
              <div className={styles.emptyState}>
                <p>No books have been created yet.</p>
              </div>
            )}
          </aside>

          <div className={styles.content}>
            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Book Details</h2>
                  <p>Title, description, and who should get submission notifications.</p>
                </div>
                <button
                  type="button"
                  className={styles.primaryButton}
                  onClick={() => void handleSaveBook()}
                  disabled={isBookSaving}
                >
                  {isBookSaving ? 'Saving...' : selectedBook ? 'Save Book' : 'Create Book'}
                </button>
              </div>

              <div className={styles.formGrid}>
                <label className={styles.field}>
                  <span>Book Title</span>
                  <input
                    type="text"
                    value={bookForm.title}
                    onChange={(event) => setBookForm((current) => ({ ...current, title: event.target.value }))}
                  />
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Description</span>
                  <textarea
                    rows={3}
                    value={bookForm.description}
                    onChange={(event) => setBookForm((current) => ({ ...current, description: event.target.value }))}
                  />
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>Admin Notification Emails</span>
                  <input
                    type="text"
                    value={bookForm.adminEmails}
                    placeholder="admin@example.com, editor@example.com"
                    onChange={(event) => setBookForm((current) => ({ ...current, adminEmails: event.target.value }))}
                  />
                </label>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Versions</h2>
                  <p>Upload source PDFs, define sections and fields, and keep one active version per book.</p>
                </div>
                <div className={styles.cardHeaderActions}>
                  {selectedBook ? (
                    <button type="button" className={styles.secondaryButton} onClick={handleCreateNewVersion}>
                      New Version
                    </button>
                  ) : null}
                  <button
                    type="button"
                    className={styles.primaryButton}
                    onClick={() => void handleSaveVersion()}
                    disabled={isVersionSaving || !selectedBook}
                  >
                    {isVersionSaving
                      ? 'Saving...'
                      : isCreatingNewVersion || !versionForm.id
                        ? 'Create Version'
                        : 'Save Version'}
                  </button>
                </div>
              </div>

              {selectedBook?.versions.length ? (
                <div className={styles.versionPills}>
                  {selectedBook.versions.map((version) => (
                    <button
                      key={version.id}
                      type="button"
                      className={`${styles.versionPill} ${selectedVersionId === version.id && !isCreatingNewVersion ? styles.versionPillActive : ''}`}
                      onClick={() => {
                        setIsCreatingNewVersion(false)
                        setSelectedVersionId(version.id)
                      }}
                    >
                      <strong>V{version.versionNumber}</strong>
                      <span>{version.isActive ? 'Active' : 'CMS only'}</span>
                      <span>{version.pendingSubmissionCount} pending</span>
                    </button>
                  ))}
                </div>
              ) : null}

              {!selectedBook ? (
                <div className={styles.emptyState}>
                  <p>Create a book before configuring versions.</p>
                </div>
              ) : isVersionLoading ? (
                <div className={styles.loaderPanel}>
                  <Loader />
                </div>
              ) : (
                <div className={styles.stack}>
                  <div className={styles.formGrid}>
                    <label className={styles.field}>
                      <span>Source Page Count</span>
                      <input
                        type="number"
                        min="1"
                        value={versionForm.sourcePageCount}
                        onChange={(event) => updateVersionForm('sourcePageCount', event.target.value)}
                      />
                    </label>

                    <label className={styles.field}>
                      <span>Content Template Page</span>
                      <input
                        type="number"
                        min="1"
                        value={versionForm.contentTemplatePageNumber}
                        onChange={(event) =>
                          updateVersionForm('contentTemplatePageNumber', event.target.value)
                        }
                      />
                    </label>

                    <label className={styles.field}>
                      <span>Section Divider Template Page</span>
                      <input
                        type="number"
                        min="1"
                        value={versionForm.sectionTemplatePageNumber}
                        onChange={(event) =>
                          updateVersionForm('sectionTemplatePageNumber', event.target.value)
                        }
                      />
                    </label>

                    <label className={styles.fileField}>
                      <span>Source PDF {isCreatingNewVersion || !versionForm.id ? '(Required)' : '(Optional replace)'}</span>
                      <input
                        type="file"
                        accept="application/pdf"
                        onChange={(event) => updateVersionForm('sourcePdfFile', event.target.files?.[0] ?? null)}
                      />
                    </label>
                  </div>

                  <div className={styles.toggleRow}>
                    <label className={styles.checkboxLabel}>
                      <input
                        type="checkbox"
                        checked={versionForm.allowPageImage}
                        onChange={(event) => updateVersionForm('allowPageImage', event.target.checked)}
                      />
                      Allow optional page images
                    </label>

                    <label className={styles.checkboxLabel}>
                      <input
                        type="checkbox"
                        checked={versionForm.allowNewSections}
                        onChange={(event) => updateVersionForm('allowNewSections', event.target.checked)}
                      />
                      Allow website requests for new sections
                    </label>

                    <label className={styles.checkboxLabel}>
                      <input
                        type="checkbox"
                        checked={versionForm.activateImmediately}
                        onChange={(event) => updateVersionForm('activateImmediately', event.target.checked)}
                      />
                      Activate this version immediately
                    </label>
                  </div>

                  <div className={styles.dualPanels}>
                    <div className={styles.inlineCard}>
                      <div className={styles.inlineCardHeader}>
                        <h3>Sections</h3>
                        <button
                          type="button"
                          className={styles.ghostButton}
                          onClick={() =>
                            updateVersionForm('sections', [
                              ...versionForm.sections,
                              { name: '', sourceStartPage: '', sourceEndPage: '' },
                            ])
                          }
                        >
                          Add Section
                        </button>
                      </div>

                      <div className={styles.collection}>
                        {versionForm.sections.map((section, index) => (
                          <div key={`section-${section.id ?? index}`} className={styles.collectionRow}>
                            <input
                              type="text"
                              placeholder="Section name"
                              value={section.name}
                              onChange={(event) =>
                                updateSection(index, { ...section, name: event.target.value })
                              }
                            />
                            <input
                              type="number"
                              min="1"
                              placeholder="Start page"
                              value={section.sourceStartPage}
                              onChange={(event) =>
                                updateSection(index, {
                                  ...section,
                                  sourceStartPage: event.target.value,
                                })
                              }
                            />
                            <input
                              type="number"
                              min="1"
                              placeholder="End page"
                              value={section.sourceEndPage}
                              onChange={(event) =>
                                updateSection(index, {
                                  ...section,
                                  sourceEndPage: event.target.value,
                                })
                              }
                            />
                            <button
                              type="button"
                              className={styles.removeButton}
                              onClick={() => removeSection(index)}
                            >
                              Remove
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className={styles.inlineCard}>
                      <div className={styles.inlineCardHeader}>
                        <h3>Submission Fields</h3>
                        <button
                          type="button"
                          className={styles.ghostButton}
                          onClick={() =>
                            updateVersionForm('fields', [
                              ...versionForm.fields,
                              {
                                label: '',
                                inputType: 'single_line',
                                placement: 'body',
                                showLabel: true,
                                isRequired: false,
                                isEmailField: false,
                              },
                            ])
                          }
                        >
                          Add Field
                        </button>
                      </div>

                      <div className={styles.fieldCollection}>
                        {versionForm.fields.map((field, index) => (
                          <div key={`field-${field.id ?? index}`} className={styles.fieldCollectionRow}>
                            <input
                              type="text"
                              placeholder="Label"
                              value={field.label}
                              onChange={(event) =>
                                updateField(index, { ...field, label: event.target.value })
                              }
                            />

                            <select
                              value={field.inputType}
                              onChange={(event) =>
                                updateField(index, {
                                  ...field,
                                  inputType: event.target.value as BookFieldInputType,
                                })
                              }
                            >
                              <option value="single_line">Single line</option>
                              <option value="rich_text">Rich text</option>
                            </select>

                            <select
                              value={field.placement}
                              onChange={(event) =>
                                updateField(index, {
                                  ...field,
                                  placement: event.target.value as BookFieldPlacement,
                                })
                              }
                            >
                              <option value="heading">Heading</option>
                              <option value="body">Body</option>
                            </select>

                            <label className={styles.inlineCheckbox}>
                              <input
                                type="checkbox"
                                checked={field.showLabel}
                                onChange={(event) =>
                                  updateField(index, { ...field, showLabel: event.target.checked })
                                }
                              />
                              Label
                            </label>

                            <label className={styles.inlineCheckbox}>
                              <input
                                type="checkbox"
                                checked={field.isRequired}
                                onChange={(event) =>
                                  updateField(index, { ...field, isRequired: event.target.checked })
                                }
                              />
                              Required
                            </label>

                            <label className={styles.inlineCheckbox}>
                              <input
                                type="checkbox"
                                checked={field.isEmailField}
                                onChange={(event) =>
                                  updateField(index, { ...field, isEmailField: event.target.checked })
                                }
                              />
                              Email field
                            </label>

                            <button
                              type="button"
                              className={styles.removeButton}
                              onClick={() => removeField(index)}
                            >
                              Remove
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  <label className={`${styles.field} ${styles.fieldFull}`}>
                    <span>Layout Settings JSON</span>
                    <textarea
                      className={styles.layoutEditor}
                      rows={18}
                      value={versionForm.layoutSettingsText}
                      onChange={(event) => updateVersionForm('layoutSettingsText', event.target.value)}
                    />
                  </label>

                  <div className={styles.versionMetaRow}>
                    <div className={styles.versionMeta}>
                      <span>
                        {selectedVersionSummary?.isActive ? 'Active on website' : 'CMS-only version'}
                      </span>
                      {selectedVersion ? (
                        <span>
                          Current section ranges already account for approved pages and new sections.
                        </span>
                      ) : null}
                    </div>

                    <div className={styles.cardHeaderActions}>
                      {selectedVersion && !selectedVersion.isActive ? (
                        <button
                          type="button"
                          className={styles.secondaryButton}
                          onClick={() => void handleActivateVersion(selectedVersion.id)}
                        >
                          Make Active
                        </button>
                      ) : null}
                      {selectedVersion ? (
                        <button
                          type="button"
                          className={styles.primaryButton}
                          onClick={() => void handleGeneratePdf()}
                          disabled={isGeneratingPdf}
                        >
                          {isGeneratingPdf ? 'Generating...' : 'Generate PDF'}
                        </button>
                      ) : null}
                    </div>
                  </div>
                </div>
              )}
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Submission Review</h2>
                  <p>Edit survivor submissions, approve them, or reject with a reason when needed.</p>
                </div>

                <div className={styles.filterGroup}>
                  <label className={styles.filterField}>
                    <span>Status</span>
                    <select
                      value={submissionStatusFilter}
                      onChange={(event) => setSubmissionStatusFilter(event.target.value)}
                    >
                      <option value="">All</option>
                      <option value="pending">Pending</option>
                      <option value="approved">Approved</option>
                      <option value="rejected">Rejected</option>
                    </select>
                  </label>
                </div>
              </div>

              {!selectedVersion ? (
                <div className={styles.emptyState}>
                  <p>Select a version to review submissions.</p>
                </div>
              ) : isSubmissionsLoading || isVersionLoading ? (
                <div className={styles.loaderPanel}>
                  <Loader />
                </div>
              ) : submissions.length > 0 ? (
                <div className={styles.submissionList}>
                  {submissions.map((submission) => (
                    <SubmissionCard
                      key={`${submission.id}-${submission.updatedAt}`}
                      submission={submission}
                      versionDetail={selectedVersion}
                      busy={busySubmissionId === submission.id}
                      onSave={handleSaveSubmission}
                      onApprove={handleApproveSubmission}
                      onReject={handleRejectSubmission}
                    />
                  ))}
                </div>
              ) : (
                <div className={styles.emptyState}>
                  <p>No submissions match the current filter.</p>
                </div>
              )}
            </section>
          </div>
        </div>
      </div>
    </CmsAppShell>
  )

  function updateVersionForm<K extends keyof VersionFormState>(key: K, value: VersionFormState[K]) {
    setVersionForm((current) => ({ ...current, [key]: value }))
  }

  function updateSection(index: number, nextSection: EditableSection) {
    setVersionForm((current) => ({
      ...current,
      sections: current.sections.map((item, itemIndex) => (itemIndex === index ? nextSection : item)),
    }))
  }

  function removeSection(index: number) {
    setVersionForm((current) => ({
      ...current,
      sections: current.sections.filter((_, itemIndex) => itemIndex !== index),
    }))
  }

  function updateField(index: number, nextField: EditableField) {
    setVersionForm((current) => ({
      ...current,
      fields: current.fields.map((item, itemIndex) => (itemIndex === index ? nextField : item)),
    }))
  }

  function removeField(index: number) {
    setVersionForm((current) => ({
      ...current,
      fields: current.fields.filter((_, itemIndex) => itemIndex !== index),
    }))
  }
}

function SubmissionCard({
  submission,
  versionDetail,
  busy,
  onSave,
  onApprove,
  onReject,
}: SubmissionCardProps) {
  const [targetMode, setTargetMode] = useState<'existing' | 'new'>(
    submission.targetSectionId ? 'existing' : 'new',
  )
  const [targetSectionId, setTargetSectionId] = useState<number | ''>(submission.targetSectionId ?? '')
  const [newSectionName, setNewSectionName] = useState(submission.newSectionName)
  const [fieldValues, setFieldValues] = useState<Record<number, string>>(
    Object.fromEntries(submission.fieldValues.map((value) => [value.fieldId, value.value])),
  )
  const [imageFile, setImageFile] = useState<File | null>(null)
  const [removeImage, setRemoveImage] = useState(false)
  const [rejectionReason, setRejectionReason] = useState(submission.rejectionReason)

  const orderedFields = useMemo(
    () => [...versionDetail.fields].sort((left, right) => left.sortOrder - right.sortOrder || left.id - right.id),
    [versionDetail.fields],
  )

  const existingTargetName =
    versionDetail.sections.find((section) => section.id === submission.targetSectionId)?.name ??
    submission.targetSectionName

  const payload = useMemo<BookSubmissionSaveInput>(
    () => ({
      targetSectionId:
        targetMode === 'existing' &&
        typeof targetSectionId === 'number' &&
        Number.isFinite(targetSectionId)
          ? targetSectionId
          : undefined,
      newSectionName: targetMode === 'new' ? newSectionName.trim() : '',
      removeImage,
      fieldValues: orderedFields.map((field) => ({
        fieldId: field.id,
        value: fieldValues[field.id] ?? '',
      })),
    }),
    [fieldValues, newSectionName, orderedFields, removeImage, targetMode, targetSectionId],
  )

  return (
    <article className={styles.submissionCard}>
      <div className={styles.submissionHeader}>
        <div>
          <h3>Submission #{submission.id}</h3>
          <p>
            Current status: <strong>{submission.status}</strong>
          </p>
          <p>
            {submission.submitterEmail
              ? `Submitter email: ${submission.submitterEmail}`
              : 'Submitter email not provided.'}
          </p>
        </div>

        <div className={styles.statusPillRow}>
          <span className={styles.statusPill}>{submission.targetSectionName || existingTargetName || 'No section'}</span>
          <span className={styles.statusPill}>{new Date(submission.createdAt).toLocaleString()}</span>
        </div>
      </div>

      <div className={styles.toggleRow}>
        <label className={styles.checkboxLabel}>
          <input
            type="radio"
            name={`target-mode-${submission.id}`}
            checked={targetMode === 'existing'}
            onChange={() => setTargetMode('existing')}
          />
          Existing section
        </label>

        {versionDetail.allowNewSections ? (
          <label className={styles.checkboxLabel}>
            <input
              type="radio"
              name={`target-mode-${submission.id}`}
              checked={targetMode === 'new'}
              onChange={() => setTargetMode('new')}
            />
            New section
          </label>
        ) : null}
      </div>

      {targetMode === 'existing' ? (
        <label className={styles.field}>
          <span>Target Section</span>
          <select
            value={typeof targetSectionId === 'number' ? String(targetSectionId) : ''}
            onChange={(event) =>
              setTargetSectionId(
                event.target.value ? Number.parseInt(event.target.value, 10) : '',
              )
            }
          >
            <option value="">Select section</option>
            {versionDetail.sections.map((section) => (
              <option key={section.id} value={section.id}>
                {section.name}
              </option>
            ))}
          </select>
        </label>
      ) : (
        <label className={styles.field}>
          <span>New Section Name</span>
          <input
            type="text"
            value={newSectionName}
            onChange={(event) => setNewSectionName(event.target.value)}
          />
        </label>
      )}

      <div className={styles.submissionFields}>
        {orderedFields.map((field) => (
          <div key={field.id} className={styles.field}>
            <span>
              {field.label}
              {field.isRequired ? ' *' : ''}
            </span>
            {field.inputType === 'rich_text' ? (
              <RichTextEditor
                value={fieldValues[field.id] ?? ''}
                onChange={(value) =>
                  setFieldValues((current) => ({ ...current, [field.id]: value }))
                }
              />
            ) : (
              <input
                type="text"
                value={fieldValues[field.id] ?? ''}
                onChange={(event) =>
                  setFieldValues((current) => ({ ...current, [field.id]: event.target.value }))
                }
              />
            )}
          </div>
        ))}
      </div>

      {versionDetail.allowPageImage ? (
        <div className={styles.imagePanel}>
          <div>
            <strong>Optional image</strong>
            <p>
              {submission.image?.fileName
                ? `Current image: ${submission.image.fileName}`
                : 'No image attached.'}
            </p>
          </div>
          <div className={styles.imagePanelControls}>
            <input
              type="file"
              accept="image/*"
              onChange={(event: ChangeEvent<HTMLInputElement>) =>
                setImageFile(event.target.files?.[0] ?? null)
              }
            />
            <label className={styles.checkboxLabel}>
              <input
                type="checkbox"
                checked={removeImage}
                onChange={(event) => setRemoveImage(event.target.checked)}
              />
              Remove current image
            </label>
          </div>
        </div>
      ) : null}

      <label className={styles.field}>
        <span>Rejection Reason</span>
        <textarea
          rows={3}
          value={rejectionReason}
          onChange={(event) => setRejectionReason(event.target.value)}
        />
      </label>

      <div className={styles.cardHeaderActions}>
        <button
          type="button"
          className={styles.secondaryButton}
          disabled={busy}
          onClick={() => void onSave(submission.id, payload, imageFile)}
        >
          {busy ? 'Working...' : 'Save Changes'}
        </button>

        {submission.status !== 'approved' ? (
          <button
            type="button"
            className={styles.primaryButton}
            disabled={busy}
            onClick={() =>
              void (async () => {
                await onSave(submission.id, payload, imageFile)
                await onApprove(submission.id)
              })()
            }
          >
            {busy ? 'Working...' : 'Approve'}
          </button>
        ) : null}

        <button
          type="button"
          className={styles.dangerButton}
          disabled={busy || !rejectionReason.trim()}
          onClick={() => void onReject(submission.id, rejectionReason)}
        >
          {busy ? 'Working...' : 'Reject'}
        </button>
      </div>
    </article>
  )
}

function buildEmptyVersionForm(detail?: BookVersionDetail | null): VersionFormState {
  if (!detail) {
    return {
      sourcePageCount: '',
      contentTemplatePageNumber: '',
      sectionTemplatePageNumber: '',
      allowPageImage: true,
      allowNewSections: true,
      activateImmediately: false,
      layoutSettingsText: JSON.stringify(DEFAULT_LAYOUT_SETTINGS, null, 2),
      sections: [{ name: '', sourceStartPage: '', sourceEndPage: '' }],
      fields: [
        {
          label: '',
          inputType: 'single_line',
          placement: 'heading',
          showLabel: false,
          isRequired: true,
          isEmailField: false,
        },
      ],
      sourcePdfFile: null,
    }
  }

  return {
    ...versionDetailToForm(detail),
    id: undefined,
    activateImmediately: false,
    sourcePdfFile: null,
  }
}

function versionDetailToForm(detail: BookVersionDetail): VersionFormState {
  return {
    id: detail.id,
    sourcePageCount: String(detail.sourcePageCount),
    contentTemplatePageNumber: String(detail.contentTemplatePageNumber),
    sectionTemplatePageNumber: String(detail.sectionTemplatePageNumber),
    allowPageImage: detail.allowPageImage,
    allowNewSections: detail.allowNewSections,
    activateImmediately: detail.isActive,
    layoutSettingsText: JSON.stringify(detail.layoutSettings ?? DEFAULT_LAYOUT_SETTINGS, null, 2),
    sections: detail.sections.map((section) => ({
      id: section.id,
      name: section.name,
      sourceStartPage: section.sourceStartPage ? String(section.sourceStartPage) : '',
      sourceEndPage: section.sourceEndPage ? String(section.sourceEndPage) : '',
    })),
    fields: detail.fields.map((field) => ({
      id: field.id,
      label: field.label,
      inputType: field.inputType,
      placement: field.placement,
      showLabel: field.showLabel,
      isRequired: field.isRequired,
      isEmailField: field.isEmailField,
    })),
    sourcePdfFile: null,
  }
}

function buildVersionPayload(form: VersionFormState): BookVersionSaveInput | null {
  try {
    const layoutSettings = JSON.parse(form.layoutSettingsText)
    return {
      sourcePageCount: Number.parseInt(form.sourcePageCount, 10),
      contentTemplatePageNumber: Number.parseInt(form.contentTemplatePageNumber, 10),
      sectionTemplatePageNumber: Number.parseInt(form.sectionTemplatePageNumber, 10),
      allowPageImage: form.allowPageImage,
      allowNewSections: form.allowNewSections,
      activateImmediately: form.activateImmediately,
      layoutSettings,
      sections: form.sections.map((section) => ({
        id: section.id,
        name: section.name.trim(),
        sourceStartPage: parseOptionalInteger(section.sourceStartPage),
        sourceEndPage: parseOptionalInteger(section.sourceEndPage),
      })),
      fields: form.fields.map((field) => ({
        id: field.id,
        label: field.label.trim(),
        inputType: field.inputType,
        placement: field.placement,
        showLabel: field.showLabel,
        isRequired: field.isRequired,
        isEmailField: field.isEmailField,
      })),
    }
  } catch {
    toast.error('Layout settings must be valid JSON.')
    return null
  }
}

async function buildSourcePdfCloneFile(
  bookId: number,
  versionDetail: BookVersionDetail | null,
) {
  if (!versionDetail) {
    return null
  }

  const sourceBlob = await booksApi.fetchSourcePdfBlob(bookId, versionDetail.id)
  return new File(
    [sourceBlob],
    `book-version-${versionDetail.versionNumber}-source.pdf`,
    { type: 'application/pdf' },
  )
}

function parseOptionalInteger(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return undefined
  }
  const parsed = Number.parseInt(trimmed, 10)
  return Number.isNaN(parsed) ? undefined : parsed
}

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}
