import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import {
  addPressMedia,
  deletePressMedia,
  fetchPressCoverImageContent,
  fetchPressEntry,
  getPressMediaContent,
} from '../api/pressApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { DateInput } from '../components/cms/DateInput'
import { EntryActions } from '../components/cms/EntryActions'
import { PublishingControls } from '../components/cms/PublishingControls'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { AddPhotoIcon, DownloadIcon, PagesIcon } from '../components/icons'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { API_BASE_URL } from '../constants/api'
import { usePressCategories } from '../hooks/usePressCategories'
import { usePressEntries } from '../hooks/usePressEntries'
import type {
  PressEntry,
  PressFormErrors,
  PressFormState,
  PressStatus,
} from '../lib/pressTypes'
import styles from '../styles/PressEditorPage.module.css'

const PRESS_MEDIA_ACCEPT = [
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

type PendingPressMedia = {
  id: string
  file: File
}

type PressMediaListItem = {
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

function emptyFormState(): PressFormState {
  return {
    title: '',
    releaseDate: '',
    categoryId: '',
    sourceUrl: '',
    contentHtml: '',
    visibility: 'public',
    publishAt: '',
    media: [],
  }
}

function entryToFormState(entry: PressEntry): PressFormState {
  return {
    title: entry.title,
    releaseDate: normalizeDateInputValue(entry.releaseDate),
    categoryId: entry.categoryId ?? '',
    sourceUrl: entry.sourceUrl,
    contentHtml: entry.contentHtml,
    visibility: entry.visibility,
    publishAt: entry.publishAt ?? '',
    media: entry.media ?? [],
  }
}

function validate(state: PressFormState, t: (k: string) => string): PressFormErrors {
  const errors: PressFormErrors = {}
  if (!state.title.trim()) {
    errors.title = t('press.editor.validation.titleRequired')
  }
  if (!state.releaseDate) {
    errors.releaseDate = t('press.editor.validation.releaseDateRequired')
  }
  if (state.sourceUrl) {
    try {
      new URL(state.sourceUrl)
    } catch {
      errors.sourceUrl = t('press.editor.validation.sourceUrlInvalid')
    }
  }
  return errors
}

type PressEditorPageProps = {
  mode?: 'create' | 'edit'
}

export function PressEditorPage({ mode = 'create' }: PressEditorPageProps) {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()
  const { create, update, remove } = usePressEntries()
  const { categories } = usePressCategories()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentEntry, setCurrentEntry] = useState<PressEntry | undefined>(undefined)
  const [isLoadingEntry, setIsLoadingEntry] = useState(isEditMode)
  const [form, setForm] = useState<PressFormState>(emptyFormState())
  const [coverImageFile, setCoverImageFile] = useState<File | undefined>()
  const [removeCoverImage, setRemoveCoverImage] = useState(false)
  const [coverPreviewUrl, setCoverPreviewUrl] = useState<string | null>(null)
  const [pendingMedia, setPendingMedia] = useState<PendingPressMedia[]>([])
  const [documentPreviewUrls, setDocumentPreviewUrls] = useState<Record<string, string>>({})
  const [activeDocumentAction, setActiveDocumentAction] = useState<
    Record<string, 'preview' | 'download' | undefined>
  >({})
  const documentPreviewUrlsRef = useRef<Record<string, string>>({})
  const [errors, setErrors] = useState<PressFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)

  useEffect(() => {
    if (!isEditMode || !id) {
      return
    }

    const loadEntry = async () => {
      try {
        setIsLoadingEntry(true)
        const entry = await fetchPressEntry(id)
        setCurrentEntry(entry)
      } catch (err) {
        toast.error(err instanceof Error ? err.message : t('press.editor.loadError'))
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
    setCoverImageFile(undefined)
    setRemoveCoverImage(false)
    setPendingMedia([])
  }, [currentEntry, isEditMode])

  useEffect(() => {
    if (coverImageFile) {
      const objectUrl = URL.createObjectURL(coverImageFile)
  
      setCoverPreviewUrl(objectUrl)
      return () => {
        URL.revokeObjectURL(objectUrl)
      }
    }

    if (removeCoverImage || !currentEntry?.coverImageUrl) {
      setCoverPreviewUrl(null)
      return
    }

    const directUrl = resolveDirectPressAssetUrl(currentEntry.coverImageUrl)
    if (directUrl) {
      setCoverPreviewUrl(directUrl)
      return
    }

    const entryId = currentEntry.id
    let cancelled = false
    let objectUrlToRevoke: string | null = null

    async function loadCoverPreview() {
      try {
        const blob = await fetchPressCoverImageContent(entryId)
        const objectUrl = URL.createObjectURL(blob)
        objectUrlToRevoke = objectUrl

        if (cancelled) {
          URL.revokeObjectURL(objectUrl)
          return
        }

        setCoverPreviewUrl(objectUrl)
      } catch {
        if (!cancelled) {
          setCoverPreviewUrl(null)
        }
      }
    }

    void loadCoverPreview()

    return () => {
      cancelled = true
      if (objectUrlToRevoke) {
        URL.revokeObjectURL(objectUrlToRevoke)
      }
    }
  }, [coverImageFile, currentEntry?.coverImageUrl, currentEntry?.id, removeCoverImage])

  const mediaItems = useMemo<PressMediaListItem[]>(
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

            const blob = await getPressMediaContent(currentEntry.id, source.id)
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

  const categoryOptions = useMemo(() => {
    if (!form.categoryId || categories.some((category) => category.id === form.categoryId)) {
      return categories
    }

    return [
      ...categories,
      {
        id: form.categoryId,
        name: t('press.editor.fields.savedCategory', { id: form.categoryId }),
        createdAt: '',
      },
    ]
  }, [categories, form.categoryId, t])

  function updateField<K extends keyof PressFormState>(key: K, value: PressFormState[K]) {
    setForm((current) => ({ ...current, [key]: value }))
    if (errors[key]) {
      setErrors((current) => {
        const next = { ...current }
        delete next[key]
        return next
      })
    }
  }

  function buildPersistedShape(state: PressFormState, status: PressStatus) {
    const publishAt = status === 'scheduled' && state.publishAt ? state.publishAt : null
    return {
      title: state.title.trim(),
      releaseDate: normalizeDateInputValue(state.releaseDate),
      categoryId: state.categoryId || null,
      sourceUrl: state.sourceUrl.trim(),
      contentHtml: state.contentHtml,
      visibility: state.visibility,
      publishAt,
      status,
    }
  }

  async function uploadPendingMedia(entryId: string, mediaEntries: PendingPressMedia[]) {
    if (!mediaEntries.length) {
      return
    }

    await addPressMedia(
      entryId,
      mediaEntries.map((item) => item.file),
      mediaEntries.map((item) => ({
        display_name: item.file.name,
        file_name: item.file.name,
      })),
    )
  }

  async function handleSave(targetStatus: PressStatus) {
    const validationErrors = validate(form, t)
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('press.editor.validation.summary'))
      return
    }

    setErrors({})
    setIsSubmitting(true)
    try {
      const payload = buildPersistedShape(form, targetStatus)
      const pendingUploads = [...pendingMedia]

      if (isEditMode && currentEntry) {
        let nextEntry = await update(currentEntry.id, payload, {
          coverImageFile,
          removeCoverImage,
        })
        await uploadPendingMedia(nextEntry.id, pendingUploads)
        if (pendingUploads.length > 0) {
          nextEntry = await fetchPressEntry(nextEntry.id)
        }
        setCurrentEntry(nextEntry)
        toast.success(
          targetStatus === 'published'
            ? t('press.feedback.published')
            : t('press.feedback.saved'),
        )
      } else {
        const created = await create(payload, { coverImageFile })
        await uploadPendingMedia(created.id, pendingUploads)
        toast.success(
          targetStatus === 'published'
            ? t('press.feedback.published')
            : t('press.feedback.saved'),
        )
        navigate(`/press/${created.id}/edit`, { replace: true })
      }
    } catch (err) {
      toast.error(err instanceof Error ? err.message : t('press.feedback.saveError'))
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
      toast.success(t('press.feedback.deleted'))
      navigate('/press', { replace: true })
    } catch (err) {
      toast.error(err instanceof Error ? err.message : t('press.feedback.deleteError'))
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
      await deletePressMedia(currentEntry.id, [numericMediaID])
      const refreshed = await fetchPressEntry(currentEntry.id)
      setCurrentEntry(refreshed)
    } catch (err) {
      toast.error(err instanceof Error ? err.message : t('press.feedback.deleteError'))
    }
  }

  function handleCoverImageSelect(files: File[]) {
    const [file] = files
    if (!file) {
      return
    }
    setCoverImageFile(file)
    setRemoveCoverImage(false)
  }

  function handleRemoveSelectedCover() {
    setCoverImageFile(undefined)
  }

  function handleRemoveExistingCover() {
    setCoverImageFile(undefined)
    setRemoveCoverImage(true)
  }

  async function handlePreviewDocument(mediaItem: PressMediaListItem) {
    setActiveDocumentAction((current) => ({
      ...current,
      [mediaItem.id]: 'preview',
    }))

    try {
      const previewUrl =
        documentPreviewUrls[mediaItem.id] ||
        (await createTemporaryPressMediaObjectUrl(currentEntry?.id, mediaItem))

      if (!previewUrl) {
        throw new Error('Document preview unavailable')
      }

      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!documentPreviewUrls[mediaItem.id]) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error(t('press.feedback.documentPreviewFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [mediaItem.id]: undefined,
      }))
    }
  }

  async function handleDownloadDocument(mediaItem: PressMediaListItem) {
    setActiveDocumentAction((current) => ({
      ...current,
      [mediaItem.id]: 'download',
    }))

    try {
      const downloadUrl = await createTemporaryPressMediaObjectUrl(currentEntry?.id, mediaItem)
      if (!downloadUrl) {
        throw new Error('Document download unavailable')
      }

      triggerFileDownload(downloadUrl, mediaItem.fileName)
      if (!documentPreviewUrls[mediaItem.id]) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error(t('press.feedback.documentDownloadFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [mediaItem.id]: undefined,
      }))
    }
  }

  if (isLoadingEntry) {
    return (
      <CmsAppShell activeKey="press">
        <div className={styles.loaderWrap}>
          <Loader label={t('press.editor.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  const currentStatus: PressStatus = currentEntry?.status ?? 'draft'
  const errorMessages = Array.from(new Set(Object.values(errors).filter(Boolean)))
  const wasModified =
    currentEntry !== undefined && currentEntry.updatedAt !== currentEntry.createdAt
  const currentCoverName =
    coverImageFile?.name || inferStoredFileName(currentEntry?.coverImageUrl, 'cover-image')

  return (
    <CmsAppShell activeKey="press">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('press.breadcrumb.entries'), to: '/press' },
            {
              label: isEditMode
                ? t('press.breadcrumb.edit')
                : t('press.breadcrumb.create'),
            },
          ]}
        />

        <div className={styles.header}>
          <div className={styles.headerText}>
            <h1>
              {isEditMode
                ? t('press.editor.titleEdit')
                : t('press.editor.titleCreate')}
            </h1>
            <p>{t('press.editor.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/press')}
          >
            {t('press.editor.backToList')}
          </button>
        </div>

        {errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p>{t('press.editor.validation.summary')}</p>
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
                <h2>{t('press.editor.sections.details')}</h2>
              </div>

              <div className={styles.fieldGrid}>
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>{t('press.editor.fields.title')}</span>
                  <input
                    type="text"
                    value={form.title}
                    placeholder={t('press.editor.fields.titlePlaceholder')}
                    onChange={(event) => updateField('title', event.target.value)}
                  />
                  {errors.title && <p className={styles.fieldError}>{errors.title}</p>}
                </label>

                <label className={styles.field}>
                  <span>{t('press.editor.fields.releaseDate')}</span>
                  <DateInput
                    value={form.releaseDate}
                    onChange={(value) => updateField('releaseDate', value)}
                  />
                  {errors.releaseDate && (
                    <p className={styles.fieldError}>{errors.releaseDate}</p>
                  )}
                </label>

                <label className={styles.field}>
                  <span>{t('press.editor.fields.category')}</span>
                  <select
                    value={form.categoryId}
                    onChange={(event) => updateField('categoryId', event.target.value)}
                  >
                    <option value="">
                      {t('press.editor.fields.categoryPlaceholder')}
                    </option>
                    {categoryOptions.map((category) => (
                      <option key={category.id} value={category.id}>
                        {category.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>{t('press.editor.fields.sourceUrl')}</span>
                  <input
                    type="url"
                    value={form.sourceUrl}
                    placeholder={t('press.editor.fields.sourceUrlPlaceholder')}
                    onChange={(event) => updateField('sourceUrl', event.target.value)}
                  />
                  {errors.sourceUrl && (
                    <p className={styles.fieldError}>{errors.sourceUrl}</p>
                  )}
                </label>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('press.editor.sections.coverImage')}</h2>
              </div>
              <UploadDropzone
                accept="image/*"
                icon={<AddPhotoIcon />}
                label={t('press.editor.media.coverLabel')}
                hint={t('press.editor.media.coverHint')}
                onFiles={handleCoverImageSelect}
              />

              {coverPreviewUrl && (
                <div className={styles.coverPreviewCard}>
                  <img
                    src={coverPreviewUrl}
                    alt={form.title || t('press.editor.media.coverImageAlt')}
                    className={styles.coverPreviewImage}
                  />
                  <div className={styles.coverPreviewContent}>
                    <h3 className={styles.coverPreviewTitle}>{currentCoverName}</h3>
                    <p className={styles.coverPreviewMeta}>
                      {coverImageFile
                        ? t('press.editor.media.pendingUpload')
                        : t('press.editor.media.currentCover')}
                    </p>
                    <div className={styles.coverPreviewActions}>
                      {coverImageFile ? (
                        <button
                          type="button"
                          className={styles.actionButtonDanger}
                          onClick={handleRemoveSelectedCover}
                        >
                          {t('press.editor.media.remove')}
                        </button>
                      ) : (
                        <button
                          type="button"
                          className={styles.actionButtonDanger}
                          onClick={handleRemoveExistingCover}
                        >
                          {t('press.editor.media.removeImage')}
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('press.editor.sections.content')}</h2>
              </div>
              <RichTextEditor
                value={form.contentHtml}
                onChange={(html) => updateField('contentHtml', html)}
                placeholder={t('press.editor.fields.contentPlaceholder')}
              />
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('press.editor.sections.media')}</h2>
                <span>{t('press.editor.sections.mediaLimit')}</span>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  multiple
                  accept={PRESS_MEDIA_ACCEPT}
                  icon={<AddPhotoIcon />}
                  label={t('press.editor.media.dropLabel')}
                  hint={t('press.editor.media.dropHint')}
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
                                  aria-label={t('press.editor.media.remove')}
                                  onClick={() => void handleRemoveMedia(mediaItem.id)}
                                >
                                  {t('press.editor.media.remove')}
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
              entryType="press"
              status={currentStatus}
              publishLabel={t('press.editor.actions.publish')}
              saveDraftLabel={t('press.editor.actions.saveDraft')}
              deleteLabel={t('press.editor.actions.delete')}
              onSaveDraft={() => void handleSave('draft')}
              onPublish={() => void handleSave('published')}
              isSubmitting={isSubmitting}
              isDeleting={isDeleting}
              onDelete={isEditMode ? () => void handleDelete() : undefined}
              deleteConfirmTitle={t('press.list.deleteDialogTitle')}
              deleteConfirmBody={t('press.editor.deleteConfirmBody', {
                title: form.title || t('press.list.untitled'),
              })}
              deleteConfirmLabel={t('press.list.delete')}
            />
          </div>

          <aside className={styles.sideColumn}>
            <PublishingControls status={currentStatus} />

            {wasModified && currentEntry && (
              <div className={styles.metaCard}>
                <span>
                  {t('press.editor.meta.created')}{' '}
                  <strong>
                    {dateFormatter.format(new Date(currentEntry.createdAt))}
                  </strong>
                </span>
                <span>
                  {t('press.editor.meta.updated')}{' '}
                  <strong>
                    {dateFormatter.format(new Date(currentEntry.updatedAt))}
                  </strong>
                </span>
              </div>
            )}

            <div className={styles.proTip}>
              <h3>{t('press.editor.proTip.title')}</h3>
              <p>{t('press.editor.proTip.body')}</p>
              <button type="button" className={styles.proTipLink}>
                {t('press.editor.proTip.cta')}
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
    isPendingUpload ? t('press.editor.media.pendingUploadShort') : '',
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

async function createTemporaryPressMediaObjectUrl(
  entryId: string | undefined,
  mediaItem: PressMediaListItem,
) {
  if (mediaItem.file) {
    return URL.createObjectURL(mediaItem.file)
  }

  if (!entryId) {
    throw new Error('Document URL unavailable')
  }

  const blob = await getPressMediaContent(entryId, mediaItem.id)
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

function resolveDirectPressAssetUrl(value?: string | null) {
  const trimmed = value?.trim() ?? ''
  if (!trimmed) {
    return ''
  }

  if (trimmed.startsWith('blob:') || trimmed.startsWith('data:')) {
    return trimmed
  }

  if (looksLikeManagedStorageReference(trimmed)) {
    return ''
  }

  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed) || trimmed.startsWith('//')) {
    return trimmed
  }

  try {
    return new URL(trimmed, `${API_BASE_URL}/`).toString()
  } catch {
    return trimmed
  }
}

function looksLikeManagedStorageReference(value: string) {
  const normalized = value.trim().toLowerCase()
  return (
    normalized.startsWith('gs://') ||
    normalized.startsWith('https://storage.googleapis.com/') ||
    normalized.startsWith('http://storage.googleapis.com/') ||
    normalized.includes('.storage.googleapis.com/')
  )
}

function inferStoredFileName(value?: string | null, fallback = 'file') {
  const trimmed = value?.trim() ?? ''
  if (!trimmed) {
    return fallback
  }

  const sanitized = trimmed.replace(/[?#].*$/, '')
  const segments = sanitized.split('/').filter(Boolean)
  return segments[segments.length - 1] || fallback
}
