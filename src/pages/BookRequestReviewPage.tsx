import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useParams } from 'react-router-dom'
import {
  booksApi,
  type BookDetail,
  type BookSubmission,
  type BookSubmissionSaveInput,
  type BookVersionDetail,
} from '../api/booksApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { generateBookPdf } from '../lib/bookPdf'
import styles from '../styles/BookRequestReviewPage.module.css'

type SubmissionValuesState = Record<number, string>

export function BookRequestReviewPage() {
  const navigate = useNavigate()
  const { bookId, submissionId } = useParams()
  const parsedBookId = bookId ? Number.parseInt(bookId, 10) : Number.NaN
  const parsedSubmissionId = submissionId ? Number.parseInt(submissionId, 10) : Number.NaN

  const [book, setBook] = useState<BookDetail | null>(null)
  const [submission, setSubmission] = useState<BookSubmission | null>(null)
  const [version, setVersion] = useState<BookVersionDetail | null>(null)
  const [targetMode, setTargetMode] = useState<'existing' | 'new'>('existing')
  const [targetSectionId, setTargetSectionId] = useState<number | ''>('')
  const [newSectionName, setNewSectionName] = useState('')
  const [fieldValues, setFieldValues] = useState<SubmissionValuesState>({})
  const [imageFile, setImageFile] = useState<File | null>(null)
  const [removeImage, setRemoveImage] = useState(false)
  const [imagePreviewUrl, setImagePreviewUrl] = useState<string | null>(null)
  const [rejectionReason, setRejectionReason] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [isWorking, setIsWorking] = useState(false)

  const orderedFields = useMemo(
    () =>
      version
        ? [...version.fields].sort((left, right) => left.sortOrder - right.sortOrder || left.id - right.id)
        : [],
    [version],
  )
  const isReadonly = !submission || submission.status !== 'pending' || isWorking

  useEffect(() => {
    if (Number.isNaN(parsedBookId) || Number.isNaN(parsedSubmissionId)) {
      setIsLoading(false)
      return
    }

    void loadRequest(parsedBookId, parsedSubmissionId)
  }, [parsedBookId, parsedSubmissionId])

  useEffect(() => {
    let cancelled = false
    let objectUrl: string | null = null

    async function loadImagePreview() {
      if (removeImage) {
        setImagePreviewUrl(null)
        return
      }

      if (imageFile) {
        objectUrl = URL.createObjectURL(imageFile)
        if (!cancelled) {
          setImagePreviewUrl(objectUrl)
        }
        return
      }

      if (!submission?.image?.fetchUrl) {
        setImagePreviewUrl(null)
        return
      }

      try {
        const blob = await booksApi.fetchSubmissionImage(submission.image.fetchUrl)
        if (cancelled) {
          return
        }
        objectUrl = URL.createObjectURL(blob)
        setImagePreviewUrl(objectUrl)
      } catch {
        if (!cancelled) {
          setImagePreviewUrl(null)
        }
      }
    }

    void loadImagePreview()

    return () => {
      cancelled = true
      if (objectUrl) {
        URL.revokeObjectURL(objectUrl)
      }
    }
  }, [imageFile, removeImage, submission?.image?.fetchUrl])

  async function loadRequest(nextBookId: number, nextSubmissionId: number) {
    try {
      setIsLoading(true)
      const [bookDetail, submissionDetail] = await Promise.all([
        booksApi.getBook(nextBookId),
        booksApi.getSubmission(nextBookId, nextSubmissionId),
      ])
      const versionDetail = await booksApi.getVersion(nextBookId, submissionDetail.bookVersionId)

      setBook(bookDetail)
      setSubmission(submissionDetail)
      setVersion(versionDetail)
      primeFormState(submissionDetail, versionDetail)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load the request.')
    } finally {
      setIsLoading(false)
    }
  }

  async function handleSave() {
    if (!book || !submission) {
      return
    }
    if (!validateSubmissionForm()) {
      return
    }

    try {
      setIsWorking(true)
      await booksApi.updateSubmission(book.id, submission.id, buildSubmissionPayload(), imageFile)
      await loadRequest(book.id, submission.id)
      toast.success('Request changes saved.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to save request changes.')
    } finally {
      setIsWorking(false)
    }
  }

  async function handleApprove() {
    if (!book || !submission) {
      return
    }
    if (!validateSubmissionForm()) {
      return
    }

    try {
      setIsWorking(true)
      await booksApi.updateSubmission(book.id, submission.id, buildSubmissionPayload(), imageFile)
      const approved = await booksApi.approveSubmission(book.id, submission.id)
      if (!approved.bookVersionId) {
        throw new Error('Approval did not return the created version.')
      }

      const createdVersion = await booksApi.getVersion(book.id, approved.bookVersionId)

      try {
        await generateAndUploadVersionPdf(book.id, book.title, createdVersion)
        toast.success(
          `Request approved. Version ${approved.bookVersionNumber ?? createdVersion.versionNumber} is now active.`,
        )
      } catch {
        toast.error(
          'Request was approved and activated, but the PDF could not be regenerated automatically. Open the book page and run Generate PDF.',
        )
      }

      navigate(`/books/${book.id}`)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to approve the request.')
    } finally {
      setIsWorking(false)
    }
  }

  async function handleReject() {
    if (!book || !submission) {
      return
    }

    if (!rejectionReason.trim()) {
      toast.error('A rejection reason is required.')
      return
    }

    try {
      setIsWorking(true)
      await booksApi.rejectSubmission(book.id, submission.id, rejectionReason.trim())
      toast.success('Request rejected.')
      navigate('/books?tab=requests')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to reject the request.')
    } finally {
      setIsWorking(false)
    }
  }

  function primeFormState(submissionDetail: BookSubmission, versionDetail: BookVersionDetail) {
    setTargetMode(submissionDetail.targetSectionId ? 'existing' : 'new')
    setTargetSectionId(submissionDetail.targetSectionId ?? '')
    setNewSectionName(submissionDetail.newSectionName)
    setFieldValues(
      Object.fromEntries(
        versionDetail.fields.map((field) => [
          field.id,
          submissionDetail.fieldValues.find((value) => value.fieldId === field.id)?.value ?? '',
        ]),
      ),
    )
    setImageFile(null)
    setRemoveImage(false)
    setRejectionReason(submissionDetail.rejectionReason)
  }

  function validateSubmissionForm() {
    if (!version) {
      return false
    }

    const missingField = orderedFields.find(
      (field) => field.isRequired && !stripHtml(fieldValues[field.id] ?? '').trim(),
    )
    if (missingField) {
      toast.error(`${missingField.label} is required.`)
      return false
    }

    if (targetMode === 'existing' && typeof targetSectionId !== 'number') {
      toast.error('Please choose the section that should receive this page.')
      return false
    }

    if (targetMode === 'new' && !version.allowNewSections) {
      toast.error('This version does not allow new sections.')
      return false
    }

    if (targetMode === 'new' && !newSectionName.trim()) {
      toast.error('Please enter the new section name.')
      return false
    }

    return true
  }

  function buildSubmissionPayload(): BookSubmissionSaveInput {
    return {
      targetSectionId:
        targetMode === 'existing' && typeof targetSectionId === 'number' ? targetSectionId : undefined,
      newSectionName: targetMode === 'new' ? newSectionName.trim() : '',
      removeImage,
      fieldValues: orderedFields.map((field) => ({
        fieldId: field.id,
        value: fieldValues[field.id] ?? '',
      })),
    }
  }

  const currentTargetLabel = useMemo(() => {
    if (!version) {
      return 'No section selected'
    }

    if (targetMode === 'existing' && typeof targetSectionId === 'number') {
      return version.sections.find((section) => section.id === targetSectionId)?.name ?? 'No section selected'
    }

    if (targetMode === 'new') {
      return newSectionName.trim() || 'New section'
    }

    return submission?.targetSectionName || submission?.newSectionName || 'No section selected'
  }, [newSectionName, submission?.newSectionName, submission?.targetSectionName, targetMode, targetSectionId, version])

  return (
    <CmsAppShell activeKey="books">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: 'Books', to: '/books' },
            { label: 'Open Requests', to: '/books?tab=requests' },
            { label: `Request #${submission?.id ?? parsedSubmissionId}` },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>{submission ? `Request #${submission.id}` : 'Request Review'}</h1>
            <p className={styles.subtitle}>
              Reviewers can correct the section target, submitted content, and optional image before
              approval. Approval creates the next version and makes it active right away.
            </p>
          </div>

          <div className={styles.headerActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/books?tab=requests')}>
              Back to Requests
            </button>
            {book ? (
              <button type="button" className={styles.secondaryButton} onClick={() => navigate(`/books/${book.id}`)}>
                Open Book
              </button>
            ) : null}
          </div>
        </header>

        {isLoading ? (
          <div className={styles.loaderWrap}>
            <Loader />
          </div>
        ) : !book || !submission || !version ? (
          <section className={styles.errorPanel}>
            <h2>Request not found</h2>
            <p>The selected request could not be loaded.</p>
          </section>
        ) : (
          <>
            <section className={styles.summaryGrid}>
              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Book</span>
                <strong>{book.title}</strong>
                <p>{book.description || 'No description added yet.'}</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Submitted Against</span>
                <strong>Version {submission.bookVersionNumber}</strong>
                <p>{currentTargetLabel}</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Submitter</span>
                <strong>{submission.submitterEmail || 'Not provided'}</strong>
                <p>{formatDateTime(submission.createdAt)}</p>
              </article>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Request Content</h2>
                  <p>Edit what will be approved into the next active version.</p>
                </div>

                <span className={styles.statusBadge}>
                  {submission.status === 'pending'
                    ? 'Pending'
                    : submission.status === 'approved'
                      ? 'Approved'
                      : 'Rejected'}
                </span>
              </div>

              <div className={styles.metaRow}>
                <span className={styles.metaPill}>Version {submission.bookVersionNumber}</span>
                <span className={styles.metaPill}>{currentTargetLabel}</span>
                {version.allowNewSections ? <span className={styles.metaTag}>New sections allowed</span> : null}
              </div>

              <div className={styles.toggleRow}>
                <label className={styles.radioLabel}>
                  <input
                    type="radio"
                    name={`target-mode-${submission.id}`}
                    checked={targetMode === 'existing'}
                    disabled={isReadonly}
                    onChange={() => setTargetMode('existing')}
                  />
                  Existing section
                </label>

                {version.allowNewSections ? (
                  <label className={styles.radioLabel}>
                    <input
                      type="radio"
                      name={`target-mode-${submission.id}`}
                      checked={targetMode === 'new'}
                      disabled={isReadonly}
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
                    disabled={isReadonly}
                    onChange={(event) =>
                      setTargetSectionId(event.target.value ? Number.parseInt(event.target.value, 10) : '')
                    }
                  >
                    <option value="">Select section</option>
                    {version.sections.map((section) => (
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
                    disabled={isReadonly}
                    onChange={(event) => setNewSectionName(event.target.value)}
                  />
                </label>
              )}

              <div className={styles.fieldStack}>
                {orderedFields.map((field) => (
                  <div key={field.id} className={styles.requestField}>
                    <div className={styles.requestFieldHeader}>
                      <div>
                        <strong>{field.label}</strong>
                        <span>
                          {field.inputType === 'rich_text' ? 'Rich text' : 'Single line'} /{' '}
                          {field.placement === 'heading' ? 'Heading' : 'Body'}
                        </span>
                      </div>

                      <div className={styles.badgeRow}>
                        {field.isRequired ? <span className={styles.metaTag}>Required</span> : null}
                        {field.isEmailField ? <span className={styles.metaTag}>Email field</span> : null}
                      </div>
                    </div>

                    {field.inputType === 'rich_text' ? (
                      <RichTextEditor
                        value={fieldValues[field.id] ?? ''}
                        disabled={isReadonly}
                        onChange={(value) =>
                          setFieldValues((current) => ({ ...current, [field.id]: value }))
                        }
                      />
                    ) : (
                      <input
                        className={styles.textInput}
                        type="text"
                        value={fieldValues[field.id] ?? ''}
                        disabled={isReadonly}
                        onChange={(event) =>
                          setFieldValues((current) => ({ ...current, [field.id]: event.target.value }))
                        }
                      />
                    )}
                  </div>
                ))}
              </div>

              {version.allowPageImage ? (
                <div className={styles.imagePanel}>
                  <div>
                    <h3>Optional Image</h3>
                    <p>
                      {submission.image?.fileName
                        ? `Current image: ${submission.image.fileName}`
                        : 'No image attached.'}
                    </p>
                  </div>

                  <div className={styles.imageControls}>
                    <input
                      type="file"
                      accept="image/*"
                      disabled={isReadonly}
                      onChange={(event) => setImageFile(event.target.files?.[0] ?? null)}
                    />
                    <label className={styles.checkboxLabel}>
                      <input
                        type="checkbox"
                        checked={removeImage}
                        disabled={isReadonly}
                        onChange={(event) => setRemoveImage(event.target.checked)}
                      />
                      Remove current image
                    </label>
                  </div>

                  {imagePreviewUrl ? (
                    <img className={styles.previewImage} src={imagePreviewUrl} alt={submission.image?.fileName || 'Preview'} />
                  ) : (
                    <div className={styles.imagePlaceholder}>No image selected</div>
                  )}
                </div>
              ) : null}
            </section>

            <section className={styles.actionCard}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Review Decision</h2>
                  <p>Rejecting needs a reason. Approving publishes the next active version for this book.</p>
                </div>
              </div>

              <label className={styles.field}>
                <span>Rejection Reason</span>
                <textarea
                  rows={4}
                  value={rejectionReason}
                  disabled={isWorking}
                  onChange={(event) => setRejectionReason(event.target.value)}
                  placeholder="Add the reason if this request should be rejected"
                />
              </label>

              <div className={styles.actionRow}>
                <button
                  type="button"
                  className={styles.secondaryButton}
                  onClick={() => void handleSave()}
                  disabled={isReadonly}
                >
                  {isWorking ? 'Working...' : 'Save Changes'}
                </button>
                <button
                  type="button"
                  className={styles.primaryButton}
                  onClick={() => void handleApprove()}
                  disabled={isReadonly}
                >
                  {isWorking ? 'Working...' : 'Approve and Publish'}
                </button>
                <button
                  type="button"
                  className={styles.dangerButton}
                  onClick={() => void handleReject()}
                  disabled={isWorking || submission.status !== 'pending' || !rejectionReason.trim()}
                >
                  {isWorking ? 'Working...' : 'Reject Request'}
                </button>
              </div>
            </section>
          </>
        )}
      </div>
    </CmsAppShell>
  )
}

async function generateAndUploadVersionPdf(bookId: number, bookTitle: string, version: BookVersionDetail) {
  const sourceBlob = await booksApi.fetchSourcePdfBlob(bookId, version.id)
  const sourceBytes = new Uint8Array(await sourceBlob.arrayBuffer())
  const generatedBlob = await generateBookPdf({
    version,
    sourcePdfBytes: sourceBytes,
    fetchImageBytes: async (submission) => {
      if (!submission.image?.fetchUrl) {
        return null
      }
      const blob = await booksApi.fetchSubmissionImage(submission.image.fetchUrl)
      return new Uint8Array(await blob.arrayBuffer())
    },
  })

  const fileName = `${slugify(bookTitle)}-version-${version.versionNumber}.pdf`
  await booksApi.uploadGeneratedPdf(bookId, version.id, generatedBlob, fileName)
}

function stripHtml(value: string) {
  return value.replace(/<[^>]*>/g, ' ')
}

function formatDateTime(value: string) {
  const parsed = Date.parse(value)
  if (Number.isNaN(parsed)) {
    return value
  }

  return new Intl.DateTimeFormat('en-CA', {
    month: 'short',
    day: '2-digit',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(parsed))
}

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}
