import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { addPressMedia, deletePressMedia, fetchPressEntry } from '../api/pressApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { PublishingControls } from '../components/cms/PublishingControls'
import { EntryActions } from '../components/cms/EntryActions'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { AddPhotoIcon } from '../components/icons'
import { usePressCategories } from '../hooks/usePressCategories'
import { usePressEntries } from '../hooks/usePressEntries'
import type {
  PressEntry,
  PressFormErrors,
  PressFormState,
  PressStatus,
} from '../lib/pressTypes'
import styles from '../styles/PressEditorPage.module.css'

type PendingPressMedia = {
  id: string
  file: File
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
    releaseDate: entry.releaseDate,
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
  const { categories, create: createCategory } = usePressCategories()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentEntry, setCurrentEntry] = useState<PressEntry | undefined>(undefined)
  const [isLoadingEntry, setIsLoadingEntry] = useState(isEditMode)

  // Fetch entry from API if in edit mode
  useEffect(() => {
    if (!isEditMode || !id) return

    const loadEntry = async () => {
      try {
        setIsLoadingEntry(true)
        const entry = await fetchPressEntry(id)
        setCurrentEntry(entry)
      } catch (err) {
        toast.error(
          err instanceof Error
            ? err.message
            : t('press.editor.loadError'),
        )
      } finally {
        setIsLoadingEntry(false)
      }
    }

    void loadEntry()
  }, [isEditMode, id, t])

  const [form, setForm] = useState<PressFormState>(
    currentEntry ? entryToFormState(currentEntry) : emptyFormState(),
  )
  const [coverImageFile, setCoverImageFile] = useState<File | undefined>()
  const [pendingMedia, setPendingMedia] = useState<PendingPressMedia[]>([])
  const [errors, setErrors] = useState<PressFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)

  useEffect(() => {
    if (isEditMode && !currentEntry) {
      return
    }

    setForm(currentEntry ? entryToFormState(currentEntry) : emptyFormState())
    setCoverImageFile(undefined)
    setPendingMedia([])
  }, [currentEntry, isEditMode])

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

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
      releaseDate: state.releaseDate,
      categoryId: state.categoryId || null,
      sourceUrl: state.sourceUrl.trim(),
      contentHtml: state.contentHtml,
      visibility: state.visibility,
      publishAt,
      status,
    }
  }

  async function uploadPendingMedia(entryId: string, mediaItems: PendingPressMedia[]) {
    if (!mediaItems.length) {
      return
    }

    await addPressMedia(
      entryId,
      mediaItems.map((item) => item.file),
      mediaItems.map((item) => ({
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
        let nextEntry = await update(currentEntry.id, payload, coverImageFile)
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
        const created = await create(payload, coverImageFile)
        await uploadPendingMedia(created.id, pendingUploads)
        toast.success(
          targetStatus === 'published'
            ? t('press.feedback.published')
            : t('press.feedback.saved'),
        )
        navigate(`/press/${created.id}/edit`, { replace: true })
      }
    } catch (err) {
      toast.error(
        err instanceof Error
          ? err.message
          : t('press.feedback.saveError'),
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
      toast.success(t('press.feedback.deleted'))
      navigate('/press', { replace: true })
    } catch (err) {
      toast.error(
        err instanceof Error
          ? err.message
          : t('press.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  function handleAddCategory() {
    const name = window.prompt(t('press.editor.categoryPrompt'))
    if (!name?.trim()) {
      return
    }
    const created = createCategory(name)
    updateField('categoryId', created.id)
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
      toast.error(err instanceof Error ? err.message : 'Failed to delete media')
    }
  }

  function handleCoverImageSelect(files: File[]) {
    if (files.length > 0) {
      setCoverImageFile(files[0])
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
                  <input
                    type="date"
                    value={form.releaseDate}
                    onChange={(event) => updateField('releaseDate', event.target.value)}
                  />
                  {errors.releaseDate && (
                    <p className={styles.fieldError}>{errors.releaseDate}</p>
                  )}
                </label>

                <label className={styles.field}>
                  <span>{t('press.editor.fields.category')}</span>
                  <div className={styles.categoryRow}>
                    <select
                      value={form.categoryId}
                      onChange={(event) => updateField('categoryId', event.target.value)}
                    >
                      <option value="">
                        {t('press.editor.fields.categoryPlaceholder')}
                      </option>
                      {categories.map((category) => (
                        <option key={category.id} value={category.id}>
                          {category.name}
                        </option>
                      ))}
                    </select>
                    <button
                      type="button"
                      className={styles.addCategoryButton}
                      onClick={handleAddCategory}
                    >
                      {t('press.editor.fields.addCategory')}
                    </button>
                  </div>
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
              {coverImageFile && (
                <div className={styles.selectedCoverImage}>
                  <span>{coverImageFile.name}</span>
                  <button
                    type="button"
                    onClick={() => setCoverImageFile(undefined)}
                  >
                    {t('press.editor.media.remove')}
                  </button>
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
                  accept="image/*,application/pdf"
                  icon={<AddPhotoIcon />}
                  label={t('press.editor.media.dropLabel')}
                  hint={t('press.editor.media.dropHint')}
                  onFiles={handleAddMedia}
                />

                {(form.media.length > 0 || pendingMedia.length > 0) && (
                  <div className={styles.uploadedList}>
                    {form.media.map((mediaItem) => (
                      <div key={mediaItem.id} className={styles.uploadedItem}>
                        <span className={styles.uploadedName}>
                          <FileIcon />
                          {mediaItem.fileName}
                        </span>
                        <button
                          type="button"
                          className={styles.removeMediaButton}
                          aria-label={t('press.editor.media.remove')}
                          onClick={() => void handleRemoveMedia(mediaItem.id)}
                        >
                          <CloseIcon />
                        </button>
                      </div>
                    ))}
                    {pendingMedia.map((mediaItem) => (
                      <div key={mediaItem.id} className={styles.uploadedItem}>
                        <span className={styles.uploadedName}>
                          <FileIcon />
                          {mediaItem.file.name}
                        </span>
                        <button
                          type="button"
                          className={styles.removeMediaButton}
                          aria-label={t('press.editor.media.remove')}
                          onClick={() => void handleRemoveMedia(mediaItem.id)}
                        >
                          <CloseIcon />
                        </button>
                      </div>
                    ))}
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
            <PublishingControls
              status={currentStatus}
            />

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

function FileIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M9 2H4.6c-.88 0-1.6.72-1.6 1.6v8.8c0 .88.72 1.6 1.6 1.6h6.8c.88 0 1.6-.72 1.6-1.6V6L9 2Z"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinejoin="round"
      />
      <path d="M9 2v3.4h4" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    </svg>
  )
}

function CloseIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M4 4l8 8M12 4l-8 8"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
      />
    </svg>
  )
}
