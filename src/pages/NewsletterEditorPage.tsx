import axios from 'axios'
import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import {
  addNewsletterMedia,
  deleteNewsletterMedia,
  fetchNewsletterEntry,
  getNewsletterMediaContent,
} from '../api/newslettersApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { DateInput } from '../components/cms/DateInput'
import { EntryActions } from '../components/cms/EntryActions'
import { PublishingControls } from '../components/cms/PublishingControls'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { AddPhotoIcon, DownloadIcon, PagesIcon } from '../components/icons'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { useNewsletterEntries } from '../hooks/useNewsletterEntries'
import type {
  NewsletterCategory,
  NewsletterEntry,
  NewsletterFormErrors,
  NewsletterFormState,
  NewsletterStatus,
} from '../lib/newsletterTypes'
import styles from '../styles/NewsletterEditorPage.module.css'

const NEWSLETTER_MEDIA_ACCEPT = [
  'image/*',
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pdf',
  '.doc',
  '.docx',
  '.xls',
  '.xlsx',
  '.ppt',
  '.pptx',
].join(',')

type PendingNewsletterMedia = {
  id: string
  file: File
}

type NewsletterMediaListItem = {
  id: string
  fileName: string
  mimeType: string
  fileSize?: number
  file?: File
  isPending: boolean
}

function makePendingMediaId() {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return `pending:${crypto.randomUUID()}`
  }
  return `pending:${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function emptyFormState(): NewsletterFormState {
  return {
    title: '',
    category: '',
    sendDate: '',
    contentHtml: '',
    visibility: 'public',
    publishAt: '',
    media: [],
  }
}

function entryToFormState(entry: NewsletterEntry): NewsletterFormState {
  return {
    title: entry.title,
    category: entry.category,
    sendDate: normalizeDateInputValue(entry.sendDate),
    contentHtml: entry.contentHtml,
    visibility: entry.visibility,
    publishAt: entry.publishAt ?? '',
    media: entry.media ?? [],
  }
}

function validate(
  state: NewsletterFormState,
  t: (key: string) => string,
): NewsletterFormErrors {
  const errors: NewsletterFormErrors = {}
  if (!state.title.trim()) {
    errors.title = t('newsletters.editor.validation.titleRequired')
  }
  if (!state.sendDate) {
    errors.sendDate = t('newsletters.editor.validation.sendDateRequired')
  }
  return errors
}

type NewsletterEditorPageProps = {
  mode?: 'create' | 'edit'
}

export function NewsletterEditorPage({
  mode = 'create',
}: NewsletterEditorPageProps) {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()
  const { create, update, remove } = useNewsletterEntries()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentEntry, setCurrentEntry] = useState<NewsletterEntry | undefined>(
    undefined,
  )
  const [isLoadingEntry, setIsLoadingEntry] = useState(isEditMode)
  const [isMissingEntry, setIsMissingEntry] = useState(false)
  const [loadErrorMessage, setLoadErrorMessage] = useState<string | null>(null)
  const [form, setForm] = useState<NewsletterFormState>(emptyFormState())
  const [pendingMedia, setPendingMedia] = useState<PendingNewsletterMedia[]>([])
  const [documentPreviewUrls, setDocumentPreviewUrls] = useState<
    Record<string, string>
  >({})
  const [activeDocumentAction, setActiveDocumentAction] = useState<
    Record<string, 'preview' | 'download' | undefined>
  >({})
  const documentPreviewUrlsRef = useRef<Record<string, string>>({})
  const [errors, setErrors] = useState<NewsletterFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)

  useEffect(() => {
    if (!isEditMode || !id) {
      return
    }

    const loadEntry = async () => {
      try {
        setIsLoadingEntry(true)
        setIsMissingEntry(false)
        setLoadErrorMessage(null)
        const entry = await fetchNewsletterEntry(id)
        setCurrentEntry(entry)
      } catch (err) {
        if (axios.isAxiosError(err) && err.response?.status === 404) {
          setIsMissingEntry(true)
          return
        }

        const message =
          err instanceof Error ? err.message : t('newsletters.editor.loadError')
        setLoadErrorMessage(message)
        toast.error(message)
      } finally {
        setIsLoadingEntry(false)
      }
    }

    void loadEntry()
  }, [id, isEditMode, t])

  useEffect(() => {
    if (isEditMode && !currentEntry) {
      return
    }

    setForm(currentEntry ? entryToFormState(currentEntry) : emptyFormState())
    setPendingMedia([])
  }, [currentEntry, isEditMode])

  const mediaItems = useMemo<NewsletterMediaListItem[]>(
    () => [
      ...form.media.map((mediaItem) => ({
        id: mediaItem.id,
        fileName: mediaItem.fileName,
        mimeType: mediaItem.mimeType ?? '',
        fileSize: mediaItem.fileSize,
        isPending: false,
      })),
      ...pendingMedia.map((mediaItem) => ({
        id: mediaItem.id,
        fileName: mediaItem.file.name,
        mimeType: mediaItem.file.type,
        fileSize: mediaItem.file.size,
        file: mediaItem.file,
        isPending: true,
      })),
    ],
    [form.media, pendingMedia],
  )

  const documentPreviewSources = useMemo(
    () =>
      mediaItems
        .filter((mediaItem) => canPreviewDocument(mediaItem.mimeType, mediaItem.fileName))
        .map((mediaItem) => ({
          id: mediaItem.id,
          file: mediaItem.file,
          sourceKey: mediaItem.file
            ? [
                mediaItem.file.name,
                mediaItem.file.size,
                mediaItem.file.lastModified,
                mediaItem.file.type,
              ].join(':')
            : `${currentEntry?.id ?? ''}:${mediaItem.id}`,
        }))
        .sort((left, right) => left.id.localeCompare(right.id)),
    [currentEntry?.id, mediaItems],
  )

  const documentPreviewSignature = useMemo(
    () =>
      JSON.stringify(
        documentPreviewSources.map((source) => ({
          id: source.id,
          sourceKey: source.sourceKey,
        })),
      ),
    [documentPreviewSources],
  )

  useEffect(() => {
    let cancelled = false

    async function loadDocumentPreviews() {
      const nextEntries = await Promise.all(
        documentPreviewSources.map(async (source) => {
          try {
            if (source.file) {
              return [source.id, URL.createObjectURL(source.file)] as const
            }

            if (!currentEntry?.id) {
              return [source.id, ''] as const
            }

            const blob = await getNewsletterMediaContent(currentEntry.id, source.id)
            return [source.id, URL.createObjectURL(blob)] as const
          } catch {
            return [source.id, ''] as const
          }
        }),
      )

      if (cancelled) {
        nextEntries.forEach(([, url]) => {
          if (url) {
            URL.revokeObjectURL(url)
          }
        })
        return
      }

      setDocumentPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })

        const next = Object.fromEntries(nextEntries.filter(([, url]) => Boolean(url)))
        documentPreviewUrlsRef.current = next
        return next
      })
    }

    if (!documentPreviewSources.length) {
      setDocumentPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })
        documentPreviewUrlsRef.current = {}
        return {}
      })
      return
    }

    void loadDocumentPreviews()

    return () => {
      cancelled = true
    }
  }, [currentEntry?.id, documentPreviewSignature, documentPreviewSources])

  useEffect(() => {
    return () => {
      Object.values(documentPreviewUrlsRef.current).forEach((url) => {
        URL.revokeObjectURL(url)
      })
      documentPreviewUrlsRef.current = {}
    }
  }, [])

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  function updateField<K extends keyof NewsletterFormState>(
    key: K,
    value: NewsletterFormState[K],
  ) {
    setForm((current) => ({ ...current, [key]: value }))
    if (errors[key]) {
      setErrors((current) => {
        const next = { ...current }
        delete next[key]
        return next
      })
    }
  }

  function buildPersistedShape(
    state: NewsletterFormState,
    status: NewsletterStatus,
  ) {
    const publishAt = status === 'scheduled' && state.publishAt ? state.publishAt : null
    return {
      title: state.title.trim(),
      category: state.category,
      sendDate: normalizeDateInputValue(state.sendDate),
      contentHtml: state.contentHtml,
      visibility: state.visibility,
      publishAt,
      status,
    }
  }

  async function uploadPendingMedia(
    entryId: string,
    mediaEntries: PendingNewsletterMedia[],
  ) {
    if (!mediaEntries.length) {
      return
    }

    await addNewsletterMedia(
      entryId,
      mediaEntries.map((item) => item.file),
      mediaEntries.map((item) => ({
        display_name: item.file.name,
        file_name: item.file.name,
      })),
    )
  }

  async function handleSave(targetStatus: NewsletterStatus) {
    const validationErrors = validate(form, t)
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('newsletters.editor.validation.summary'))
      return
    }

    setErrors({})
    setIsSubmitting(true)
    try {
      const payload = buildPersistedShape(form, targetStatus)
      const pendingUploads = [...pendingMedia]

      if (isEditMode && currentEntry) {
        let nextEntry = await update(currentEntry.id, payload)
        await uploadPendingMedia(nextEntry.id, pendingUploads)
        if (pendingUploads.length > 0) {
          nextEntry = await fetchNewsletterEntry(nextEntry.id)
        }
        setCurrentEntry(nextEntry)
        toast.success(
          targetStatus === 'published'
            ? t('newsletters.feedback.published')
            : t('newsletters.feedback.saved'),
        )
      } else {
        const created = await create(payload)
        await uploadPendingMedia(created.id, pendingUploads)
        toast.success(
          targetStatus === 'published'
            ? t('newsletters.feedback.published')
            : t('newsletters.feedback.saved'),
        )
        navigate(`/newsletters/${created.id}/edit`, { replace: true })
      }
    } catch (err) {
      toast.error(
        err instanceof Error ? err.message : t('newsletters.feedback.saveError'),
      )
    } finally {
      setIsSubmitting(false)
    }
  }

  async function handleDelete() {
    if (!currentEntry) {
      return
    }
    setIsDeleting(true)
    try {
      await remove(currentEntry.id)
      toast.success(t('newsletters.feedback.deleted'))
      navigate('/newsletters', { replace: true })
    } catch (err) {
      toast.error(
        err instanceof Error ? err.message : t('newsletters.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  function handleAddMedia(files: File[]) {
    if (!files.length) {
      return
    }

    setPendingMedia((current) => [
      ...current,
      ...files.map((file) => ({
        id: makePendingMediaId(),
        file,
      })),
    ])
  }

  async function handleRemoveMedia(mediaId: string) {
    if (mediaId.startsWith('pending:')) {
      setPendingMedia((current) => current.filter((item) => item.id !== mediaId))
      return
    }

    if (!currentEntry) {
      return
    }

    const numericMediaID = Number.parseInt(mediaId, 10)
    if (Number.isNaN(numericMediaID)) {
      return
    }

    try {
      await deleteNewsletterMedia(currentEntry.id, [numericMediaID])
      const refreshed = await fetchNewsletterEntry(currentEntry.id)
      setCurrentEntry(refreshed)
    } catch (err) {
      toast.error(
        err instanceof Error ? err.message : t('newsletters.feedback.deleteError'),
      )
    }
  }

  async function handlePreviewDocument(mediaItem: NewsletterMediaListItem) {
    setActiveDocumentAction((current) => ({
      ...current,
      [mediaItem.id]: 'preview',
    }))

    try {
      const previewUrl =
        documentPreviewUrls[mediaItem.id] ||
        (await createTemporaryNewsletterMediaObjectUrl(currentEntry?.id, mediaItem))

      if (!previewUrl) {
        throw new Error('Document preview unavailable')
      }

      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!documentPreviewUrls[mediaItem.id]) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error(t('newsletters.feedback.documentPreviewFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [mediaItem.id]: undefined,
      }))
    }
  }

  async function handleDownloadDocument(mediaItem: NewsletterMediaListItem) {
    setActiveDocumentAction((current) => ({
      ...current,
      [mediaItem.id]: 'download',
    }))

    try {
      const downloadUrl = await createTemporaryNewsletterMediaObjectUrl(
        currentEntry?.id,
        mediaItem,
      )
      if (!downloadUrl) {
        throw new Error('Document download unavailable')
      }

      triggerFileDownload(downloadUrl, mediaItem.fileName)
      if (!documentPreviewUrls[mediaItem.id]) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error(t('newsletters.feedback.documentDownloadFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [mediaItem.id]: undefined,
      }))
    }
  }

  if (isLoadingEntry) {
    return (
      <CmsAppShell activeKey="newsletters">
        <div className={styles.loaderWrap}>
          <Loader label={t('newsletters.editor.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (isMissingEntry) {
    return (
      <CmsAppShell activeKey="newsletters">
        <div className={styles.page}>
          <Breadcrumb
            items={[
              { label: t('newsletters.breadcrumb.entries'), to: '/newsletters' },
              { label: t('newsletters.editor.notFoundTitle') },
            ]}
          />
          <section className={styles.card}>
            <h1>{t('newsletters.editor.notFoundTitle')}</h1>
            <p>{t('newsletters.editor.notFoundText')}</p>
            <button
              type="button"
              className={styles.backLink}
              onClick={() => navigate('/newsletters')}
            >
              {t('newsletters.editor.backToList')}
            </button>
          </section>
        </div>
      </CmsAppShell>
    )
  }

  if (isEditMode && !currentEntry && loadErrorMessage) {
    return (
      <CmsAppShell activeKey="newsletters">
        <div className={styles.page}>
          <Breadcrumb
            items={[
              { label: t('newsletters.breadcrumb.entries'), to: '/newsletters' },
              { label: t('newsletters.editor.titleEdit') },
            ]}
          />
          <section className={styles.card}>
            <h1>{t('newsletters.editor.loadError')}</h1>
            <p>{loadErrorMessage}</p>
            <button
              type="button"
              className={styles.backLink}
              onClick={() => navigate('/newsletters')}
            >
              {t('newsletters.editor.backToList')}
            </button>
          </section>
        </div>
      </CmsAppShell>
    )
  }

  const currentStatus: NewsletterStatus = currentEntry?.status ?? 'draft'
  const errorMessages = Array.from(new Set(Object.values(errors).filter(Boolean)))
  const wasModified =
    currentEntry !== undefined && currentEntry.updatedAt !== currentEntry.createdAt

  return (
    <CmsAppShell activeKey="newsletters">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('newsletters.breadcrumb.entries'), to: '/newsletters' },
            {
              label: isEditMode
                ? t('newsletters.breadcrumb.edit')
                : t('newsletters.breadcrumb.create'),
            },
          ]}
        />

        <div className={styles.header}>
          <div className={styles.headerText}>
            <h1>
              {isEditMode
                ? t('newsletters.editor.titleEdit')
                : t('newsletters.editor.titleCreate')}
            </h1>
            <p>{t('newsletters.editor.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/newsletters')}
          >
            {t('newsletters.editor.backToList')}
          </button>
        </div>

        {errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p>{t('newsletters.editor.validation.summary')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        <div className={styles.layout}>
          <div className={styles.mainColumn}>
            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('newsletters.editor.sections.details')}</h2>
              </div>

              <div className={styles.fieldGrid}>
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>{t('newsletters.editor.fields.title')}</span>
                  <input
                    type="text"
                    value={form.title}
                    placeholder={t('newsletters.editor.fields.titlePlaceholder')}
                    onChange={(event) => updateField('title', event.target.value)}
                  />
                  {errors.title && <p className={styles.fieldError}>{errors.title}</p>}
                </label>

                <label className={styles.field}>
                  <span>{t('newsletters.editor.fields.category')}</span>
                  <select
                    value={form.category}
                    onChange={(event) =>
                      updateField('category', event.target.value as NewsletterCategory)
                    }
                  >
                    <option value="">
                      {t('newsletters.editor.fields.categoryPlaceholder')}
                    </option>
                    <option value="csaa">
                      {t('newsletters.editor.fields.categoryCsaa')}
                    </option>
                    <option value="cst">
                      {t('newsletters.editor.fields.categoryCst')}
                    </option>
                  </select>
                </label>

                <label className={styles.field}>
                  <span>{t('newsletters.editor.fields.sendDate')}</span>
                  <DateInput
                    value={form.sendDate}
                    onChange={(value) => updateField('sendDate', value)}
                  />
                  {errors.sendDate && (
                    <p className={styles.fieldError}>{errors.sendDate}</p>
                  )}
                </label>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('newsletters.editor.sections.content')}</h2>
              </div>
              <RichTextEditor
                value={form.contentHtml}
                onChange={(html) => updateField('contentHtml', html)}
                placeholder={t('newsletters.editor.fields.contentPlaceholder')}
              />
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('newsletters.editor.sections.attachments')}</h2>
                <span>{t('newsletters.editor.sections.attachmentsLimit')}</span>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  multiple
                  accept={NEWSLETTER_MEDIA_ACCEPT}
                  icon={<AddPhotoIcon />}
                  label={t('newsletters.editor.media.dropLabel')}
                  hint={t('newsletters.editor.media.dropHint')}
                  onFiles={handleAddMedia}
                />

                {mediaItems.length > 0 && (
                  <div className={styles.documentList}>
                    {mediaItems.map((mediaItem) => {
                      const canPreview = canPreviewDocument(
                        mediaItem.mimeType,
                        mediaItem.fileName,
                      )
                      const previewUrl = documentPreviewUrls[mediaItem.id] || ''
                      const documentTypeLabel = resolveDocumentTypeLabel(
                        mediaItem.mimeType,
                        mediaItem.fileName,
                      )
                      const documentTitle = stripFileExtension(mediaItem.fileName)
                      const documentMeta = buildDocumentMeta({
                        isPendingUpload: mediaItem.isPending,
                        fileTypeLabel: documentTypeLabel,
                        fileSize: mediaItem.fileSize,
                        t,
                      })

                      return (
                        <article key={mediaItem.id} className={styles.documentItem}>
                          <div className={styles.documentPreviewRow}>
                            <div className={styles.documentPreviewCard} aria-hidden="true">
                              {canPreview && previewUrl ? (
                                <iframe
                                  src={previewUrl}
                                  title={documentTitle}
                                  className={styles.documentPreviewFrame}
                                />
                              ) : (
                                <>
                                  <PagesIcon
                                    size={28}
                                    className={styles.documentPreviewIcon}
                                  />
                                  <span className={styles.documentPreviewBadge}>
                                    {documentTypeLabel}
                                  </span>
                                </>
                              )}
                            </div>
                            <div className={styles.documentPreviewContent}>
                              <h4 className={styles.documentName}>{documentTitle}</h4>
                              {documentMeta ? (
                                <p className={styles.documentMeta}>{documentMeta}</p>
                              ) : null}
                              <div className={styles.documentPreviewActions}>
                                {canPreview ? (
                                  <button
                                    type="button"
                                    className={styles.actionButton}
                                    disabled={activeDocumentAction[mediaItem.id] !== undefined}
                                    onClick={() => void handlePreviewDocument(mediaItem)}
                                  >
                                    {activeDocumentAction[mediaItem.id] === 'preview'
                                      ? t('pages.common.loading')
                                      : t('pages.modules.actions.previewDocument')}
                                  </button>
                                ) : null}
                                <button
                                  type="button"
                                  className={styles.actionButton}
                                  disabled={activeDocumentAction[mediaItem.id] !== undefined}
                                  onClick={() => void handleDownloadDocument(mediaItem)}
                                >
                                  <DownloadIcon size={16} />
                                  {activeDocumentAction[mediaItem.id] === 'download'
                                    ? t('pages.common.loading')
                                    : t('pages.modules.actions.downloadDocument')}
                                </button>
                                <button
                                  type="button"
                                  className={styles.actionButtonDanger}
                                  aria-label={t('newsletters.editor.media.remove')}
                                  onClick={() => void handleRemoveMedia(mediaItem.id)}
                                >
                                  {t('newsletters.editor.media.remove')}
                                </button>
                              </div>
                            </div>
                          </div>
                        </article>
                      )
                    })}
                  </div>
                )}
              </div>
            </section>

            <EntryActions
              entryType="newsletter"
              status={currentStatus}
              publishLabel={t('newsletters.editor.actions.publish')}
              saveDraftLabel={t('newsletters.editor.actions.saveDraft')}
              deleteLabel={t('newsletters.editor.actions.delete')}
              onSaveDraft={() => void handleSave('draft')}
              onPublish={() => void handleSave('published')}
              isSubmitting={isSubmitting}
              isDeleting={isDeleting}
              onDelete={isEditMode ? () => void handleDelete() : undefined}
              deleteConfirmTitle={t('newsletters.list.deleteDialogTitle')}
              deleteConfirmBody={t('newsletters.editor.deleteConfirmBody', {
                title: form.title || t('newsletters.list.untitled'),
              })}
              deleteConfirmLabel={t('newsletters.list.delete')}
            />
          </div>

          <aside className={styles.sideColumn}>
            <PublishingControls status={currentStatus} />

            {wasModified && currentEntry && (
              <div className={styles.metaCard}>
                <span>
                  {t('newsletters.editor.meta.created')}{' '}
                  <strong>
                    {dateFormatter.format(new Date(currentEntry.createdAt))}
                  </strong>
                </span>
                <span>
                  {t('newsletters.editor.meta.updated')}{' '}
                  <strong>
                    {dateFormatter.format(new Date(currentEntry.updatedAt))}
                  </strong>
                </span>
              </div>
            )}

            <div className={styles.proTip}>
              <h3>{t('newsletters.editor.proTip.title')}</h3>
              <p>{t('newsletters.editor.proTip.body')}</p>
              <button type="button" className={styles.proTipLink}>
                {t('newsletters.editor.proTip.cta')}
              </button>
            </div>
          </aside>
        </div>
      </div>
    </CmsAppShell>
  )
}

function stripFileExtension(value: string) {
  return value.replace(/\.[^.]+$/, '')
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

function resolveDocumentTypeLabel(mimeType: string, fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (extension) {
    return extension.length > 5 ? extension.slice(0, 5) : extension
  }

  const normalizedMimeType = mimeType.trim().toLowerCase()
  if (normalizedMimeType.includes('pdf')) {
    return 'PDF'
  }
  if (normalizedMimeType.includes('word')) {
    return 'DOC'
  }
  if (normalizedMimeType.includes('excel') || normalizedMimeType.includes('sheet')) {
    return 'XLS'
  }
  if (
    normalizedMimeType.includes('powerpoint') ||
    normalizedMimeType.includes('presentation')
  ) {
    return 'PPT'
  }

  return 'FILE'
}

function buildDocumentMeta({
  isPendingUpload,
  fileTypeLabel,
  fileSize,
  t,
}: {
  isPendingUpload: boolean
  fileTypeLabel: string
  fileSize?: number
  t: (key: string) => string
}) {
  const parts = [
    isPendingUpload ? t('newsletters.editor.media.pendingUploadShort') : '',
    fileTypeLabel,
    typeof fileSize === 'number' && fileSize > 0 ? formatFileSize(fileSize) : '',
  ].filter(Boolean)

  return parts.join(' | ')
}

function canPreviewDocument(mimeType: string, fileName: string) {
  const normalizedMimeType = mimeType.trim().toLowerCase()
  const extension = fileName.split('.').pop()?.trim().toLowerCase() ?? ''

  return (
    normalizedMimeType.includes('pdf') ||
    normalizedMimeType.startsWith('image/') ||
    ['pdf', 'png', 'jpg', 'jpeg', 'webp', 'gif'].includes(extension)
  )
}

async function createTemporaryNewsletterMediaObjectUrl(
  entryId: string | undefined,
  mediaItem: NewsletterMediaListItem,
) {
  if (mediaItem.file) {
    return URL.createObjectURL(mediaItem.file)
  }

  if (!entryId) {
    throw new Error('Document URL unavailable')
  }

  const blob = await getNewsletterMediaContent(entryId, mediaItem.id)
  return URL.createObjectURL(blob)
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

function normalizeDateInputValue(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return ''
  }

  const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/)
  if (match) {
    return match[1]
  }

  const parsed = Date.parse(trimmed)
  if (Number.isNaN(parsed)) {
    return trimmed
  }

  return new Date(parsed).toISOString().slice(0, 10)
}
