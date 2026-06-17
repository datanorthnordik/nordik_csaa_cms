import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useParams } from 'react-router-dom'
import {
  booksApi,
  type BookDetail,
  type BookFieldInputType,
  type BookFieldPlacement,
  type BookVersionDetail,
  type BookVersionSaveInput,
  type BookVersionSection,
} from '../api/booksApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { Toggle } from '../components/Toggle'
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
  sourcePageCount: string
  contentTemplatePageNumber: string
  sectionTemplatePageNumber: string
  allowPageImage: boolean
  allowNewSections: boolean
  layoutSettingsText: string
  sections: EditableSection[]
  fields: EditableField[]
  sourcePdfFile: File | null
}

export function BooksPage() {
  const navigate = useNavigate()
  const { id } = useParams()
  const selectedBookId = id ? Number.parseInt(id, 10) : Number.NaN

  const [selectedBook, setSelectedBook] = useState<BookDetail | null>(null)
  const [selectedVersionId, setSelectedVersionId] = useState<number | null>(null)
  const [selectedVersion, setSelectedVersion] = useState<BookVersionDetail | null>(null)
  const [bookForm, setBookForm] = useState({
    title: '',
    description: '',
  })
  const [versionForm, setVersionForm] = useState<VersionFormState>(buildEmptyVersionForm())
  const [isBookLoading, setIsBookLoading] = useState(true)
  const [isBookSaving, setIsBookSaving] = useState(false)
  const [isVersionLoading, setIsVersionLoading] = useState(false)
  const [isVersionSaving, setIsVersionSaving] = useState(false)
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false)

  const selectedVersionSummary = useMemo(
    () => selectedBook?.versions.find((version) => version.id === selectedVersionId) ?? null,
    [selectedBook, selectedVersionId],
  )
  const activeVersionSummary = useMemo(
    () =>
      selectedBook?.activeVersionId
        ? selectedBook.versions.find((version) => version.id === selectedBook.activeVersionId) ?? null
        : null,
    [selectedBook],
  )
  const pendingSubmissionCount = useMemo(
    () => selectedBook?.versions.reduce((total, version) => total + version.pendingSubmissionCount, 0) ?? 0,
    [selectedBook],
  )
  const hasVersions = (selectedBook?.versions.length ?? 0) > 0

  useEffect(() => {
    if (Number.isNaN(selectedBookId)) {
      setSelectedBook(null)
      setSelectedVersion(null)
      setSelectedVersionId(null)
      setIsBookLoading(false)
      return
    }

    void loadBookDetail(selectedBookId, { preserveSelection: true })
  }, [selectedBookId])

  useEffect(() => {
    if (!selectedBook?.id || !selectedVersionId) {
      return
    }

    void loadVersionDetail(selectedBook.id, selectedVersionId)
  }, [selectedBook?.id, selectedVersionId])

  async function loadBookDetail(
    bookId: number,
    options: { preserveSelection?: boolean; preferredVersionId?: number | null } = {},
  ) {
    try {
      setIsBookLoading(true)
      const detail = await booksApi.getBook(bookId)
      setSelectedBook(detail)
      setBookForm({
        title: detail.title,
        description: detail.description,
      })

      if (detail.versions.length === 0) {
        setSelectedVersion(null)
        setSelectedVersionId(null)
        setVersionForm(buildEmptyVersionForm())
        return
      }

      const nextVersionId =
        options.preferredVersionId && detail.versions.some((version) => version.id === options.preferredVersionId)
          ? options.preferredVersionId
          : options.preserveSelection && selectedVersionId && detail.versions.some((version) => version.id === selectedVersionId)
            ? selectedVersionId
            : detail.activeVersionId ?? detail.versions[0]?.id ?? null

      setSelectedVersionId(nextVersionId)
      if (!nextVersionId) {
        setSelectedVersion(null)
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load the selected book.')
    } finally {
      setIsBookLoading(false)
    }
  }

  async function loadVersionDetail(bookId: number, versionId: number) {
    try {
      setIsVersionLoading(true)
      const detail = await booksApi.getVersion(bookId, versionId)
      setSelectedVersion(detail)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to load the selected version.')
    } finally {
      setIsVersionLoading(false)
    }
  }

  async function handleSaveBook() {
    if (!selectedBook?.id) {
      return
    }

    const title = bookForm.title.trim()
    const description = bookForm.description.trim()
    if (!title) {
      toast.error('Book title is required.')
      return
    }

    try {
      setIsBookSaving(true)
      await booksApi.updateBook(selectedBook.id, {
        title,
        description,
        adminNotificationEmails: selectedBook.adminNotificationEmails,
      })
      await loadBookDetail(selectedBook.id, { preserveSelection: true })
      toast.success('Book details saved.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to save book details.')
    } finally {
      setIsBookSaving(false)
    }
  }

  async function handleCreateInitialVersion() {
    if (!selectedBook?.id) {
      toast.error('Save the book before creating the initial version.')
      return
    }

    const payload = buildVersionPayload(versionForm)
    if (!payload) {
      return
    }
    if (!versionForm.sourcePdfFile) {
      toast.error('A source PDF is required for the initial version.')
      return
    }

    try {
      setIsVersionSaving(true)
      const created = await booksApi.createVersion(selectedBook.id, payload, versionForm.sourcePdfFile)
      await loadBookDetail(selectedBook.id, {
        preserveSelection: true,
        preferredVersionId: created.id,
      })
      await loadVersionDetail(selectedBook.id, created.id)
      toast.success('Initial version created and set active.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to create the initial version.')
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
      await loadBookDetail(selectedBook.id, { preserveSelection: true, preferredVersionId: versionId })
      await loadVersionDetail(selectedBook.id, versionId)
      toast.success('Version activated.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to activate the version.')
    }
  }

  async function handleGeneratePdf(versionOverride?: BookVersionDetail) {
    if (!selectedBook?.id || (!selectedVersionId && !versionOverride)) {
      return
    }

    try {
      setIsGeneratingPdf(true)
      const detail = versionOverride ?? (await booksApi.getVersion(selectedBook.id, selectedVersionId as number))
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
      await loadBookDetail(selectedBook.id, { preserveSelection: true, preferredVersionId: detail.id })
      await loadVersionDetail(selectedBook.id, detail.id)
      toast.success('Generated PDF uploaded.')
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Unable to generate the PDF.')
    } finally {
      setIsGeneratingPdf(false)
    }
  }

  return (
    <CmsAppShell activeKey="books">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: 'Books', to: '/books' },
            { label: selectedBook?.title || 'Book Workspace' },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>{selectedBook?.title || 'Book Workspace'}</h1>
            <p className={styles.subtitle}>
              Manage the one-time initial version setup here, then review later versions created from
              approved website requests and switch the active version whenever needed.
            </p>
          </div>

          <div className={styles.headerActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/books?tab=requests')}>
              Open Requests
            </button>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/books')}>
              Back to Books
            </button>
          </div>
        </header>

        {isBookLoading ? (
          <div className={styles.loaderWrap}>
            <Loader />
          </div>
        ) : !selectedBook ? (
          <section className={styles.errorPanel}>
            <h2>Book not found</h2>
            <p>The requested book could not be loaded.</p>
          </section>
        ) : (
          <>
            <section className={styles.summaryGrid}>
              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Pending Requests</span>
                <strong>{pendingSubmissionCount}</strong>
                <p>Website requests waiting for review.</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Versions</span>
                <strong>{selectedBook.versions.length}</strong>
                <p>Initial setup plus automatically created approval versions.</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Active Version</span>
                <strong>{activeVersionSummary ? `Version ${activeVersionSummary.versionNumber}` : 'None'}</strong>
                <p>{activeVersionSummary ? 'Currently visible on the website.' : 'No website version is active yet.'}</p>
              </article>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Book Details</h2>
                  <p>Update the title and description used across the CMS and website.</p>
                </div>

                <button type="button" className={styles.primaryButton} onClick={() => void handleSaveBook()} disabled={isBookSaving}>
                  {isBookSaving ? 'Saving...' : 'Save Book'}
                </button>
              </div>

              <div className={styles.fieldStack}>
                <label className={styles.field}>
                  <span>Book Title</span>
                  <input
                    type="text"
                    value={bookForm.title}
                    onChange={(event) => setBookForm((current) => ({ ...current, title: event.target.value }))}
                  />
                </label>

                <label className={styles.field}>
                  <span>Description</span>
                  <textarea
                    rows={4}
                    value={bookForm.description}
                    onChange={(event) => setBookForm((current) => ({ ...current, description: event.target.value }))}
                  />
                </label>
              </div>
            </section>

            {!hasVersions ? (
              <section className={styles.card}>
                <div className={styles.cardHeader}>
                  <div>
                    <h2>Initial Version Setup</h2>
                    <p>
                      This schema is configured once in CMS. After it is created, later versions come
                      only from approved website requests.
                    </p>
                  </div>
                </div>

                <div className={styles.setupHint}>
                  <span className={styles.metaTag}>One-time setup</span>
                  <p>
                    Configure the source PDF, template pages, section ranges, and survivor submission
                    schema here. This setup is locked after the first version is saved.
                  </p>
                </div>

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
                      onChange={(event) => updateVersionForm('contentTemplatePageNumber', event.target.value)}
                    />
                  </label>

                  <label className={styles.field}>
                    <span>Section Divider Template Page</span>
                    <input
                      type="number"
                      min="1"
                      value={versionForm.sectionTemplatePageNumber}
                      onChange={(event) => updateVersionForm('sectionTemplatePageNumber', event.target.value)}
                    />
                  </label>

                  <label className={`${styles.field} ${styles.fieldFull}`}>
                    <span>Source PDF</span>
                    <input
                      type="file"
                      accept="application/pdf"
                      onChange={(event) => updateVersionForm('sourcePdfFile', event.target.files?.[0] ?? null)}
                    />
                  </label>
                </div>

                <div className={styles.featureToggleRow}>
                  <label className={styles.featureToggle}>
                    <span>Allow optional page images</span>
                    <Toggle
                      checked={versionForm.allowPageImage}
                      onChange={(checked) => updateVersionForm('allowPageImage', checked)}
                    />
                  </label>

                  <label className={styles.featureToggle}>
                    <span>Allow website requests for new sections</span>
                    <Toggle
                      checked={versionForm.allowNewSections}
                      onChange={(checked) => updateVersionForm('allowNewSections', checked)}
                    />
                  </label>
                </div>

                <div className={styles.twoColumn}>
                  <article className={styles.inlineCard}>
                    <div className={styles.inlineCardHeader}>
                      <div>
                        <h4>Sections</h4>
                        <p>Define the named source PDF sections that website requests can target.</p>
                      </div>
                      <button type="button" className={styles.secondaryButton} onClick={addSection}>
                        Add Section
                      </button>
                    </div>

                    <div className={styles.collectionStack}>
                      {versionForm.sections.map((section, index) => (
                        <div key={`${section.id ?? 'new'}-${index}`} className={styles.collectionRow}>
                          <input
                            type="text"
                            value={section.name}
                            placeholder="Section name"
                            onChange={(event) => updateSection(index, { ...section, name: event.target.value })}
                          />
                          <input
                            type="number"
                            min="1"
                            value={section.sourceStartPage}
                            placeholder="Start"
                            onChange={(event) =>
                              updateSection(index, { ...section, sourceStartPage: event.target.value })
                            }
                          />
                          <input
                            type="number"
                            min="1"
                            value={section.sourceEndPage}
                            placeholder="End"
                            onChange={(event) =>
                              updateSection(index, { ...section, sourceEndPage: event.target.value })
                            }
                          />
                          <button type="button" className={styles.dangerButton} onClick={() => removeSection(index)}>
                            Remove
                          </button>
                        </div>
                      ))}
                    </div>
                  </article>

                  <article className={styles.inlineCard}>
                    <div className={styles.inlineCardHeader}>
                      <div>
                        <h4>Submission Fields</h4>
                        <p>Define the configurable heading and body fields contributors will submit.</p>
                      </div>
                      <button type="button" className={styles.secondaryButton} onClick={addField}>
                        Add Field
                      </button>
                    </div>

                    <div className={styles.collectionStack}>
                      {versionForm.fields.map((field, index) => (
                        <div key={`${field.id ?? 'new'}-${index}`} className={styles.fieldCollectionItem}>
                          <div className={styles.fieldRow1}>
                            <input
                              type="text"
                              value={field.label}
                              placeholder="Field label"
                              onChange={(event) => updateField(index, { ...field, label: event.target.value })}
                            />
                          </div>

                          <div className={styles.fieldRow2}>
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
                          </div>

                          <div className={styles.fieldRow3}>
                            <label className={styles.toggleLabel}>
                              <span>Show label</span>
                              <Toggle
                                checked={field.showLabel}
                                onChange={(checked) => updateField(index, { ...field, showLabel: checked })}
                              />
                            </label>

                            <label className={styles.toggleLabel}>
                              <span>Required</span>
                              <Toggle
                                checked={field.isRequired}
                                onChange={(checked) => updateField(index, { ...field, isRequired: checked })}
                              />
                            </label>

                            <label className={styles.toggleLabel}>
                              <span>Email field</span>
                              <Toggle
                                checked={field.isEmailField}
                                onChange={(checked) => updateField(index, { ...field, isEmailField: checked })}
                              />
                            </label>

                            <button type="button" className={styles.dangerButton} onClick={() => removeField(index)}>
                              Remove
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </article>
                </div>

                <div className={styles.formFooter}>
                  <button
                    type="button"
                    className={styles.primaryButton}
                    onClick={() => void handleCreateInitialVersion()}
                    disabled={isVersionSaving}
                  >
                    {isVersionSaving ? 'Creating...' : 'Create Initial Version'}
                  </button>
                </div>
              </section>
            ) : (
              <section className={styles.card}>
                <div className={styles.cardHeader}>
                  <div>
                    <h2>Versions</h2>
                    <p>
                      The initial setup is now locked. Every approved website request creates the next
                      active version automatically, and you can still switch back to earlier versions here.
                    </p>
                  </div>
                </div>

                <div className={styles.workspace}>
                  <aside className={styles.versionRail}>
                    {selectedBook.versions.map((version) => (
                      <button
                        key={version.id}
                        type="button"
                        className={[
                          styles.versionCard,
                          selectedVersionId === version.id ? styles.versionCardActive : '',
                        ].join(' ')}
                        onClick={() => setSelectedVersionId(version.id)}
                      >
                        <div className={styles.versionCardTop}>
                          <strong>Version {version.versionNumber}</strong>
                          {version.isActive ? <span className={styles.statusBadge}>Active</span> : null}
                        </div>
                        <p>{version.pendingSubmissionCount} pending request{version.pendingSubmissionCount === 1 ? '' : 's'}</p>
                        <small>Updated {formatDate(version.updatedAt)}</small>
                      </button>
                    ))}
                  </aside>

                  <div className={styles.versionPanel}>
                    {isVersionLoading ? (
                      <div className={styles.loaderWrap}>
                        <Loader />
                      </div>
                    ) : selectedVersion ? (
                      <>
                        <div className={styles.versionHeader}>
                          <div>
                            <div className={styles.versionTitleRow}>
                              <h3>Version {selectedVersion.versionNumber}</h3>
                              <span className={styles.metaPill}>
                                {selectedVersion.isActive ? 'Active on website' : 'Inactive version'}
                              </span>
                            </div>
                            <p className={styles.versionCopy}>
                              Approval creates the latest live version automatically. Use this page if
                              you want to switch the website back to a previous version later.
                            </p>
                          </div>

                          <div className={styles.actionRow}>
                            {!selectedVersion.isActive ? (
                              <button
                                type="button"
                                className={styles.secondaryButton}
                                onClick={() => void handleActivateVersion(selectedVersion.id)}
                              >
                                Make Active
                              </button>
                            ) : null}
                            <button
                              type="button"
                              className={styles.primaryButton}
                              onClick={() => void handleGeneratePdf()}
                              disabled={isGeneratingPdf}
                            >
                              {isGeneratingPdf ? 'Generating...' : 'Generate PDF'}
                            </button>
                          </div>
                        </div>

                        <div className={styles.infoGrid}>
                          <InfoTile label="Source Page Count" value={String(selectedVersion.sourcePageCount)} />
                          <InfoTile label="Content Template Page" value={String(selectedVersion.contentTemplatePageNumber)} />
                          <InfoTile label="Section Divider Page" value={String(selectedVersion.sectionTemplatePageNumber)} />
                          <InfoTile label="Approved Requests" value={String(selectedVersion.approvedSubmissions.length)} />
                        </div>

                        <div className={styles.metaRow}>
                          <span className={styles.metaTag}>
                            {selectedVersion.allowPageImage ? 'Optional page images enabled' : 'No optional page images'}
                          </span>
                          <span className={styles.metaTag}>
                            {selectedVersion.allowNewSections ? 'New sections allowed' : 'New sections disabled'}
                          </span>
                          {selectedVersionSummary?.lastGeneratedAt ? (
                            <span className={styles.metaTag}>Generated {formatDateTime(selectedVersionSummary.lastGeneratedAt)}</span>
                          ) : null}
                        </div>

                        <div className={styles.twoColumn}>
                          <article className={styles.inlineCard}>
                            <div className={styles.inlineCardHeader}>
                              <div>
                                <h4>Sections</h4>
                                <p>These ranges are fixed from the initial setup and updated by approvals.</p>
                              </div>
                            </div>

                            <div className={styles.listStack}>
                              {selectedVersion.sections.map((section) => (
                                <div key={section.id} className={styles.listRow}>
                                  <strong>{section.name}</strong>
                                  <span>{formatSectionPages(section)}</span>
                                </div>
                              ))}
                            </div>
                          </article>

                          <article className={styles.inlineCard}>
                            <div className={styles.inlineCardHeader}>
                              <div>
                                <h4>Submission Fields</h4>
                                <p>The website form for this version is based on the original CMS setup.</p>
                              </div>
                            </div>

                            <div className={styles.listStack}>
                              {selectedVersion.fields.map((field) => (
                                <div key={field.id} className={styles.fieldRow}>
                                  <div>
                                    <strong>{field.label}</strong>
                                    <span>
                                      {field.inputType === 'rich_text' ? 'Rich text' : 'Single line'} /{' '}
                                      {field.placement === 'heading' ? 'Heading' : 'Body'}
                                    </span>
                                  </div>
                                  <div className={styles.badgeRow}>
                                    {field.isRequired ? <span className={styles.smallBadge}>Required</span> : null}
                                    {field.isEmailField ? <span className={styles.smallBadge}>Email</span> : null}
                                    {!field.showLabel ? <span className={styles.smallBadgeMuted}>Label hidden</span> : null}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </article>
                        </div>
                      </>
                    ) : (
                      <div className={styles.emptyState}>
                        <h3>Select a version</h3>
                        <p>Choose a version from the left to review its locked setup and live content.</p>
                      </div>
                    )}
                  </div>
                </div>
              </section>
            )}
          </>
        )}
      </div>
    </CmsAppShell>
  )

  function updateVersionForm<K extends keyof VersionFormState>(key: K, value: VersionFormState[K]) {
    setVersionForm((current) => ({ ...current, [key]: value }))
  }

  function addSection() {
    setVersionForm((current) => ({
      ...current,
      sections: [...current.sections, { name: '', sourceStartPage: '', sourceEndPage: '' }],
    }))
  }

  function updateSection(index: number, nextSection: EditableSection) {
    setVersionForm((current) => ({
      ...current,
      sections: current.sections.map((section, itemIndex) => (itemIndex === index ? nextSection : section)),
    }))
  }

  function removeSection(index: number) {
    setVersionForm((current) => ({
      ...current,
      sections: current.sections.filter((_, itemIndex) => itemIndex !== index),
    }))
  }

  function addField() {
    setVersionForm((current) => ({
      ...current,
      fields: [
        ...current.fields,
        {
          label: '',
          inputType: 'single_line',
          placement: 'body',
          showLabel: true,
          isRequired: false,
          isEmailField: false,
        },
      ],
    }))
  }

  function updateField(index: number, nextField: EditableField) {
    setVersionForm((current) => ({
      ...current,
      fields: current.fields.map((field, itemIndex) => (itemIndex === index ? nextField : field)),
    }))
  }

  function removeField(index: number) {
    setVersionForm((current) => ({
      ...current,
      fields: current.fields.filter((_, itemIndex) => itemIndex !== index),
    }))
  }
}

function InfoTile({ label, value }: { label: string; value: string }) {
  return (
    <article className={styles.infoTile}>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  )
}

function buildEmptyVersionForm(): VersionFormState {
  return {
    sourcePageCount: '',
    contentTemplatePageNumber: '',
    sectionTemplatePageNumber: '',
    allowPageImage: true,
    allowNewSections: true,
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

function buildVersionPayload(form: VersionFormState): BookVersionSaveInput | null {
  try {
    const layoutSettings = JSON.parse(form.layoutSettingsText)
    return {
      sourcePageCount: Number.parseInt(form.sourcePageCount, 10),
      contentTemplatePageNumber: Number.parseInt(form.contentTemplatePageNumber, 10),
      sectionTemplatePageNumber: Number.parseInt(form.sectionTemplatePageNumber, 10),
      allowPageImage: form.allowPageImage,
      allowNewSections: form.allowNewSections,
      activateImmediately: true,
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

function formatSectionPages(section: BookVersionSection) {
  if (section.sourceStartPage != null && section.sourceEndPage != null) {
    return `Source pages ${section.sourceStartPage}-${section.sourceEndPage}`
  }
  return 'Generated section from approved requests'
}

function parseOptionalInteger(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return undefined
  }
  const parsed = Number.parseInt(trimmed, 10)
  return Number.isNaN(parsed) ? undefined : parsed
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
