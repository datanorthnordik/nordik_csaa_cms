import { useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { PublishingControls } from '../components/cms/PublishingControls'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { AddPhotoIcon } from '../components/icons'
import { usePressCategories } from '../hooks/usePressCategories'
import { usePressEntries } from '../hooks/usePressEntries'
import type {
  PressEntry,
  PressFormErrors,
  PressFormState,
  PressMedia,
  PressStatus,
} from '../lib/pressTypes'
import styles from '../styles/PressEditorPage.module.css'

function makeMediaId() {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }
  return `media-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

function readMediaUrl(file: File): Promise<string> {
  if (file.type.startsWith('image/')) {
    return readAsDataUrl(file)
  }
  return Promise.resolve(URL.createObjectURL(file))
}

function firstImageUrl(media: PressMedia[]): string | undefined {
  return media.find((item) => item.mimeType?.startsWith('image/'))?.fileUrl
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
  const { entries, create, update, remove, get } = usePressEntries()
  const { categories, create: createCategory } = usePressCategories()

  const isEditMode = mode === 'edit' && Boolean(id)
  const currentEntry = isEditMode && id ? get(id) : undefined
  const isMissingEntry = isEditMode && !currentEntry && entries.length > 0

  const [form, setForm] = useState<PressFormState>(() =>
    currentEntry ? entryToFormState(currentEntry) : emptyFormState(),
  )
  const [errors, setErrors] = useState<PressFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const lastHydratedKey = useRef<string | null>(currentEntry?.id ?? null)
  const currentKey = currentEntry?.id ?? (isEditMode ? null : 'new')
  if (currentKey !== lastHydratedKey.current) {
    lastHydratedKey.current = currentKey
    setForm(currentEntry ? entryToFormState(currentEntry) : emptyFormState())
  }

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
      coverImageUrl: firstImageUrl(state.media),
      media: state.media,
      status,
    }
  }

  function handleSave(targetStatus: PressStatus) {
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
      if (isEditMode && currentEntry) {
        update(currentEntry.id, payload)
        toast.success(
          targetStatus === 'published'
            ? t('press.feedback.published')
            : t('press.feedback.saved'),
        )
      } else {
        const created = create(payload)
        toast.success(
          targetStatus === 'published'
            ? t('press.feedback.published')
            : t('press.feedback.saved'),
        )
        navigate(`/press/${created.id}/edit`, { replace: true })
      }
    } finally {
      setIsSubmitting(false)
    }
  }

  function handleDelete() {
    if (!currentEntry) {
      return
    }
    setIsDeleting(true)
    try {
      remove(currentEntry.id)
      toast.success(t('press.feedback.deleted'))
      navigate('/press', { replace: true })
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

  async function handleAddMedia(files: File[]) {
    if (!files.length) {
      return
    }
    const newMedia: PressMedia[] = await Promise.all(
      files.map(async (file) => ({
        id: makeMediaId(),
        fileName: file.name,
        fileUrl: await readMediaUrl(file),
        mimeType: file.type,
        fileSize: file.size,
      })),
    )
    setForm((current) => ({ ...current, media: [...current.media, ...newMedia] }))
  }

  function handleRemoveMedia(mediaId: string) {
    setForm((current) => {
      const removed = current.media.find((m) => m.id === mediaId)
      if (removed?.fileUrl?.startsWith('blob:')) {
        URL.revokeObjectURL(removed.fileUrl)
      }
      return {
        ...current,
        media: current.media.filter((m) => m.id !== mediaId),
      }
    })
  }

  if (isMissingEntry) {
    return (
      <CmsAppShell activeKey="press">
        <div className={styles.page}>
          <Breadcrumb
            items={[
              { label: t('press.breadcrumb.entries'), to: '/press' },
              { label: t('press.editor.notFoundTitle') },
            ]}
          />
          <section className={styles.card}>
            <h1>{t('press.editor.notFoundTitle')}</h1>
            <p>{t('press.editor.notFoundText')}</p>
            <button
              type="button"
              className={styles.backLink}
              onClick={() => navigate('/press')}
            >
              {t('press.editor.backToList')}
            </button>
          </section>
        </div>
      </CmsAppShell>
    )
  }

  if (isEditMode && !currentEntry) {
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

                {form.media.length > 0 && (
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
                          onClick={() => handleRemoveMedia(mediaItem.id)}
                        >
                          <CloseIcon />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </section>
          </div>

          <aside className={styles.sideColumn}>
            <PublishingControls
              status={currentStatus}
              visibility={form.visibility}
              onVisibilityChange={(value) => updateField('visibility', value)}
              publishLabel={t('press.editor.actions.publish')}
              saveDraftLabel={t('press.editor.actions.saveDraft')}
              onSaveDraft={() => handleSave('draft')}
              onPublish={() => handleSave('published')}
              isSubmitting={isSubmitting}
              isDeleting={isDeleting}
              onDelete={isEditMode ? handleDelete : undefined}
              deleteConfirmTitle={t('press.list.deleteDialogTitle')}
              deleteConfirmBody={t('press.editor.deleteConfirmBody', {
                title: form.title || t('press.list.untitled'),
              })}
              deleteConfirmLabel={t('press.list.delete')}
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
