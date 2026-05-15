import { useMemo, useRef, useState } from 'react'
import { DateInput } from '../components/cms/DateInput'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { PublishingControls } from '../components/cms/PublishingControls'
import { EntryActions } from '../components/cms/EntryActions'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { AddPhotoIcon } from '../components/icons'
import { useNewsletterEntries } from '../hooks/useNewsletterEntries'
import type {
  NewsletterCategory,
  NewsletterEntry,
  NewsletterFormErrors,
  NewsletterFormState,
  NewsletterMedia,
  NewsletterStatus,
} from '../lib/newsletterTypes'
import styles from '../styles/NewsletterEditorPage.module.css'

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
    sendDate: entry.sendDate,
    contentHtml: entry.contentHtml,
    visibility: entry.visibility,
    publishAt: entry.publishAt ?? '',
    media: entry.media ?? [],
  }
}

function validate(
  state: NewsletterFormState,
  t: (k: string) => string,
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

export function NewsletterEditorPage({ mode = 'create' }: NewsletterEditorPageProps) {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()
  const { entries, create, update, remove, get } = useNewsletterEntries()

  const isEditMode = mode === 'edit' && Boolean(id)
  const currentEntry = isEditMode && id ? get(id) : undefined
  const isMissingEntry = isEditMode && !currentEntry && entries.length > 0

  const [form, setForm] = useState<NewsletterFormState>(() =>
    currentEntry ? entryToFormState(currentEntry) : emptyFormState(),
  )
  const [errors, setErrors] = useState<NewsletterFormErrors>({})
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

  function buildPersistedShape(state: NewsletterFormState, status: NewsletterStatus) {
    const publishAt = status === 'scheduled' && state.publishAt ? state.publishAt : null
    return {
      title: state.title.trim(),
      category: state.category,
      sendDate: state.sendDate,
      contentHtml: state.contentHtml,
      visibility: state.visibility,
      publishAt,
      media: state.media,
      status,
    }
  }

  function handleSave(targetStatus: NewsletterStatus) {
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
      if (isEditMode && currentEntry) {
        update(currentEntry.id, payload)
        toast.success(
          targetStatus === 'published'
            ? t('newsletters.feedback.published')
            : t('newsletters.feedback.saved'),
        )
      } else {
        const created = create(payload)
        toast.success(
          targetStatus === 'published'
            ? t('newsletters.feedback.published')
            : t('newsletters.feedback.saved'),
        )
        navigate(`/newsletters/${created.id}/edit`, { replace: true })
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
      toast.success(t('newsletters.feedback.deleted'))
      navigate('/newsletters', { replace: true })
    } finally {
      setIsDeleting(false)
    }
  }

  async function handleAddMedia(files: File[]) {
    if (!files.length) {
      return
    }
    const newMedia: NewsletterMedia[] = await Promise.all(
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

  if (isEditMode && !currentEntry) {
    return (
      <CmsAppShell activeKey="newsletters">
        <div className={styles.loaderWrap}>
          <Loader label={t('newsletters.editor.loading')} />
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
                  {errors.title && (
                    <p className={styles.fieldError}>{errors.title}</p>
                  )}
                </label>

                <label className={styles.field}>
                  <span>{t('newsletters.editor.fields.category')}</span>
                  <select
                    value={form.category}
                    onChange={(event) =>
                      updateField('category', event.target.value as NewsletterCategory)
                    }
                  >
                    <option value="" disabled>
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
                  accept="image/*,application/pdf"
                  icon={<AddPhotoIcon />}
                  label={t('newsletters.editor.media.dropLabel')}
                  hint={t('newsletters.editor.media.dropHint')}
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
                          aria-label={t('newsletters.editor.media.remove')}
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

            <EntryActions
              entryType="newsletter"
              status={currentStatus}
              publishLabel={t('newsletters.editor.actions.publish')}
              saveDraftLabel={t('newsletters.editor.actions.saveDraft')}
              deleteLabel={t('newsletters.editor.actions.delete')}
              onSaveDraft={() => handleSave('draft')}
              onPublish={() => handleSave('published')}
              isSubmitting={isSubmitting}
              isDeleting={isDeleting}
              onDelete={isEditMode ? handleDelete : undefined}
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
