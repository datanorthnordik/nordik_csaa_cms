import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { memorialApi } from '../api/memorialApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { EntryActions } from '../components/cms/EntryActions'
import { PublishingControls } from '../components/cms/PublishingControls'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { CloudUploadIcon } from '../components/icons'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { useMemorialEntries } from '../hooks/useMemorialEntries'
import {
  MEMORIAL_CATEGORIES,
  type MemorialCategory,
  type MemorialEntry,
  type MemorialFormErrors,
  type MemorialFormState,
  type MemorialGalleryImage,
  type MemorialStatus,
} from '../lib/memorialTypes'
import styles from '../styles/MemorialEntryEditorPage.module.css'

const MEMORIAL_IMAGE_ACCEPT = 'image/png,image/jpeg,image/webp,image/gif'

type ExistingGalleryPreview = MemorialGalleryImage & {
  previewUrl: string
}

type NewGalleryPreview = {
  localId: string
  file: File
  previewUrl: string
}

type LoadedMediaState = {
  portraitPreviewUrl: string | null
  galleryImages: ExistingGalleryPreview[]
}

type MemorialEntryEditorPageProps = {
  mode?: 'create' | 'edit'
}

function emptyForm(): MemorialFormState {
  return {
    fullName: '',
    biography: '',
    category: '',
    affiliation: '',
    status: 'draft',
    dateOfBirth: '',
    dateOfPassing: '',
  }
}

function memorialToFormState(entry: MemorialEntry): MemorialFormState {
  return {
    fullName: entry.fullName,
    biography: entry.biography,
    category: entry.category,
    affiliation: entry.affiliation,
    status: entry.status,
    dateOfBirth: entry.dateOfBirth,
    dateOfPassing: entry.dateOfPassing,
  }
}

export function MemorialEntryEditorPage({
  mode = 'create',
}: MemorialEntryEditorPageProps) {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()
  const { create, update, remove } = useMemorialEntries()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentEntry, setCurrentEntry] = useState<MemorialEntry | null>(null)
  const [isLoading, setIsLoading] = useState(isEditMode)
  const [isSaving, setIsSaving] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const [form, setForm] = useState<MemorialFormState>(emptyForm())
  const [errors, setErrors] = useState<MemorialFormErrors>({})
  const [portraitFile, setPortraitFile] = useState<File | null>(null)
  const [portraitPreviewUrl, setPortraitPreviewUrl] = useState<string | null>(null)
  const [removePortrait, setRemovePortrait] = useState(false)
  const [existingGalleryImages, setExistingGalleryImages] = useState<ExistingGalleryPreview[]>([])
  const [newGalleryImages, setNewGalleryImages] = useState<NewGalleryPreview[]>([])
  const [removedGalleryImageIds, setRemovedGalleryImageIds] = useState<number[]>([])

  const portraitPreviewRef = useRef<string | null>(null)
  const existingGalleryRef = useRef<ExistingGalleryPreview[]>([])
  const newGalleryRef = useRef<NewGalleryPreview[]>([])

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  useEffect(() => {
    return () => {
      if (portraitPreviewRef.current) {
        URL.revokeObjectURL(portraitPreviewRef.current)
      }
      for (const image of existingGalleryRef.current) {
        if (image.previewUrl) {
          URL.revokeObjectURL(image.previewUrl)
        }
      }
      for (const image of newGalleryRef.current) {
        URL.revokeObjectURL(image.previewUrl)
      }
    }
  }, [])

  useEffect(() => {
    if (!isEditMode || !id) {
      replacePortraitPreview(null)
      replaceExistingGalleryImages([])
      replaceNewGalleryImages([])
      setCurrentEntry(null)
      setForm(emptyForm())
      setRemovePortrait(false)
      setRemovedGalleryImageIds([])
      setPortraitFile(null)
      return
    }

    let cancelled = false

    const loadEntry = async () => {
      try {
        setIsLoading(true)
        const entry = await memorialApi.getMemorial(id)
        const media = await loadMediaPreviews(entry)
        if (cancelled) {
          revokeLoadedMedia(media)
          return
        }
        applyLoadedEntry(entry, media)
      } catch (loadError) {
        toast.error(
          loadError instanceof Error
            ? loadError.message
            : t('memorial.feedback.loadError'),
        )
      } finally {
        if (!cancelled) {
          setIsLoading(false)
        }
      }
    }

    void loadEntry()

    return () => {
      cancelled = true
    }
  }, [id, isEditMode, t])

  function replacePortraitPreview(nextUrl: string | null) {
    if (portraitPreviewRef.current && portraitPreviewRef.current !== nextUrl) {
      URL.revokeObjectURL(portraitPreviewRef.current)
    }
    portraitPreviewRef.current = nextUrl
    setPortraitPreviewUrl(nextUrl)
  }

  function replaceExistingGalleryImages(nextImages: ExistingGalleryPreview[]) {
    const nextUrls = new Set(
      nextImages
        .map((image) => image.previewUrl)
        .filter((value): value is string => Boolean(value)),
    )

    for (const image of existingGalleryRef.current) {
      if (image.previewUrl && !nextUrls.has(image.previewUrl)) {
        URL.revokeObjectURL(image.previewUrl)
      }
    }

    existingGalleryRef.current = nextImages
    setExistingGalleryImages(nextImages)
  }

  function replaceNewGalleryImages(nextImages: NewGalleryPreview[]) {
    const nextUrls = new Set(nextImages.map((image) => image.previewUrl))

    for (const image of newGalleryRef.current) {
      if (!nextUrls.has(image.previewUrl)) {
        URL.revokeObjectURL(image.previewUrl)
      }
    }

    newGalleryRef.current = nextImages
    setNewGalleryImages(nextImages)
  }

  async function loadMediaPreviews(entry: MemorialEntry): Promise<LoadedMediaState> {
    let nextPortraitUrl: string | null = null

    if (entry.portrait) {
      try {
        const blob = await memorialApi.getMemorialPortraitContent(entry.id)
        nextPortraitUrl = URL.createObjectURL(blob)
      } catch {
        nextPortraitUrl = null
      }
    }

    const galleryImages = await Promise.all(
      entry.galleryImages.map(async (image) => {
        try {
          const blob = await memorialApi.getMemorialGalleryImageContent(
            entry.id,
            image.id,
          )
          return {
            ...image,
            previewUrl: URL.createObjectURL(blob),
          }
        } catch {
          return {
            ...image,
            previewUrl: '',
          }
        }
      }),
    )

    return {
      portraitPreviewUrl: nextPortraitUrl,
      galleryImages,
    }
  }

  function revokeLoadedMedia(media: LoadedMediaState) {
    if (media.portraitPreviewUrl) {
      URL.revokeObjectURL(media.portraitPreviewUrl)
    }
    for (const image of media.galleryImages) {
      if (image.previewUrl) {
        URL.revokeObjectURL(image.previewUrl)
      }
    }
  }

  function applyLoadedEntry(entry: MemorialEntry, media: LoadedMediaState) {
    setCurrentEntry(entry)
    setForm(memorialToFormState(entry))
    setErrors({})
    setPortraitFile(null)
    setRemovePortrait(false)
    setRemovedGalleryImageIds([])
    replacePortraitPreview(media.portraitPreviewUrl)
    replaceExistingGalleryImages(media.galleryImages)
    replaceNewGalleryImages([])
  }

  function updateField<K extends keyof MemorialFormState>(
    key: K,
    value: MemorialFormState[K],
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

  function validate() {
    const nextErrors: MemorialFormErrors = {}

    if (!form.fullName.trim()) {
      nextErrors.fullName = t('memorial.editor.validation.fullNameRequired')
    }
    if (!form.category) {
      nextErrors.category = t('memorial.editor.validation.categoryRequired')
    }

    return nextErrors
  }

  async function handleSave(publish = false) {
    const validationErrors = validate()
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('memorial.editor.validation.summary'))
      return
    }

    const statusToPersist: MemorialStatus = publish ? 'published' : form.status

    setErrors({})
    setIsSaving(true)
    try {
      const input = {
        full_name: form.fullName.trim(),
        affiliation: form.affiliation.trim(),
        category: form.category as MemorialCategory,
        status: statusToPersist,
        biography: form.biography,
        date_of_birth: form.dateOfBirth || undefined,
        date_of_passing: form.dateOfPassing || undefined,
        remove_portrait: removePortrait,
        remove_gallery_image_ids: removedGalleryImageIds,
      }

      if (isEditMode && currentEntry) {
        const updated = await update(
          currentEntry.id,
          input,
          portraitFile,
          newGalleryImages.map((image) => image.file),
        )
        const media = await loadMediaPreviews(updated)
        applyLoadedEntry(updated, media)
        toast.success(t('memorial.feedback.updated'))
      } else {
        const created = await create(
          input,
          portraitFile,
          newGalleryImages.map((image) => image.file),
        )
        toast.success(t('memorial.feedback.created'))
        navigate(`/memorial/${created.id}/edit`, { replace: true })
      }
    } catch (saveError) {
      toast.error(
        saveError instanceof Error
          ? saveError.message
          : t('memorial.feedback.saveError'),
      )
    } finally {
      setIsSaving(false)
    }
  }

  async function handleDelete() {
    if (!currentEntry) {
      return
    }

    setIsDeleting(true)
    try {
      await remove(currentEntry.id)
      toast.success(t('memorial.feedback.deleted'))
      navigate('/memorial', { replace: true })
    } catch (deleteError) {
      toast.error(
        deleteError instanceof Error
          ? deleteError.message
          : t('memorial.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  function handlePortraitFiles(files: File[]) {
    const file = files[0]
    if (!file) {
      return
    }

    setPortraitFile(file)
    setRemovePortrait(false)
    replacePortraitPreview(URL.createObjectURL(file))
  }

  function removeSelectedPortrait() {
    setPortraitFile(null)
    setRemovePortrait(true)
    replacePortraitPreview(null)
  }

  function handleGalleryFiles(files: File[]) {
    const nextImages = files.map((file, index) => ({
      localId: `${Date.now()}-${index}-${file.name}`,
      file,
      previewUrl: URL.createObjectURL(file),
    }))
    replaceNewGalleryImages([...newGalleryImages, ...nextImages])
  }

  function removeExistingGalleryImage(imageId: string) {
    setRemovedGalleryImageIds((current) => {
      const numericId = Number(imageId)
      return current.includes(numericId) ? current : [...current, numericId]
    })
    replaceExistingGalleryImages(
      existingGalleryImages.filter((image) => image.id !== imageId),
    )
  }

  function removeNewGalleryImage(localId: string) {
    replaceNewGalleryImages(
      newGalleryImages.filter((image) => image.localId !== localId),
    )
  }

  function formatDate(value?: string) {
    if (!value) {
      return '—'
    }
    const parsed = Date.parse(value)
    if (Number.isNaN(parsed)) {
      return value
    }
    return dateFormatter.format(new Date(parsed))
  }

  const pageTitle = isEditMode
    ? t('memorial.editor.editTitle')
    : t('memorial.editor.createTitle')

  const errorMessages = Array.from(new Set(Object.values(errors).filter(Boolean)))
  const galleryItemCount = existingGalleryImages.length + newGalleryImages.length

  if (isLoading) {
    return (
      <CmsAppShell activeKey="memorial">
        <div className={styles.loaderWrap}>
          <Loader label={t('memorial.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="memorial">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            {
              label: t('memorial.breadcrumb.entries'),
              to: '/memorial',
            },
            {
              label: isEditMode
                ? currentEntry?.fullName || pageTitle
                : t('memorial.breadcrumb.createNew'),
            },
          ]}
        />

        <header className={styles.header}>
          <div className={styles.headerText}>
            <h1>{pageTitle}</h1>
            <p>{t('memorial.editor.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/memorial')}
          >
            {t('memorial.editor.backToEntries')}
          </button>
        </header>

        {errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p>{t('memorial.editor.validation.summary')}</p>
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
                <h2>{t('memorial.editor.sections.entryDetails')}</h2>
              </div>

              <div className={styles.fieldGrid}>
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.fullName')}
                  </span>
                  <input
                    type="text"
                    value={form.fullName}
                    onChange={(event) =>
                      updateField('fullName', event.target.value)
                    }
                    placeholder={t('memorial.editor.fields.fullNamePlaceholder')}
                  />
                  {errors.fullName && (
                    <p className={styles.fieldError}>{errors.fullName}</p>
                  )}
                </label>

                <label className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.category')}
                  </span>
                  <select
                    value={form.category}
                    onChange={(event) =>
                      updateField(
                        'category',
                        event.target.value as MemorialFormState['category'],
                      )
                    }
                  >
                    <option value="">
                      {t('memorial.editor.fields.categoryPlaceholder')}
                    </option>
                    {MEMORIAL_CATEGORIES.map((category) => (
                      <option key={category} value={category}>
                        {t(`memorial.category.${category}`)}
                      </option>
                    ))}
                  </select>
                  {errors.category && (
                    <p className={styles.fieldError}>{errors.category}</p>
                  )}
                </label>

                <label className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.status')}
                  </span>
                  <select
                    value={form.status}
                    onChange={(event) =>
                      updateField(
                        'status',
                        event.target.value as MemorialFormState['status'],
                      )
                    }
                  >
                    <option value="draft">{t('memorial.status.draft')}</option>
                    <option value="review">{t('memorial.status.review')}</option>
                    <option value="published">
                      {t('memorial.status.published')}
                    </option>
                  </select>
                </label>

                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.affiliation')}
                  </span>
                  <input
                    type="text"
                    value={form.affiliation}
                    onChange={(event) =>
                      updateField('affiliation', event.target.value)
                    }
                    placeholder={t('memorial.editor.fields.affiliationPlaceholder')}
                  />
                </label>
              </div>
            </section>

            <section className={styles.biographyCard}>
              <p className={styles.biographyLabel}>
                {t('memorial.editor.sections.biography')}
              </p>
              <RichTextEditor
                value={form.biography}
                onChange={(html) => updateField('biography', html)}
                placeholder={t('memorial.editor.fields.biographyPlaceholder')}
                allowImages={false}
              />
            </section>

            <section className={styles.mediaCard}>
              <div className={styles.cardHeader}>
                <h2>{t('memorial.editor.sections.media')}</h2>
              </div>

              <div className={styles.mediaLayout}>
                <div className={styles.portraitZone}>
                  <span className={styles.galleryLabel}>
                    {t('memorial.editor.sections.featuredPortrait')}
                  </span>
                  {portraitPreviewUrl ? (
                    <div className={styles.portraitPreviewCard}>
                      <div className={styles.portraitPreview}>
                        <img
                          src={portraitPreviewUrl}
                          alt={form.fullName || t('memorial.editor.sections.featuredPortrait')}
                          className={styles.portraitPreviewImg}
                        />
                      </div>
                      <div className={styles.mediaMeta}>
                        <strong>
                          {portraitFile?.name ||
                            currentEntry?.portrait?.fileName ||
                            t('memorial.editor.sections.portraitSelected')}
                        </strong>
                        <span>{t('memorial.editor.sections.featuredPortraitHint')}</span>
                      </div>
                      <div className={styles.inlineActions}>
                        <label className={styles.secondaryButton}>
                          <input
                            type="file"
                            accept={MEMORIAL_IMAGE_ACCEPT}
                            className={styles.hiddenInput}
                            onChange={(event) =>
                              handlePortraitFiles(
                                Array.from(event.target.files ?? []),
                              )
                            }
                          />
                          {t('memorial.editor.sections.replacePortrait')}
                        </label>
                        <button
                          type="button"
                          className={styles.removeButton}
                          onClick={removeSelectedPortrait}
                        >
                          {t('memorial.editor.sections.removePortrait')}
                        </button>
                      </div>
                    </div>
                  ) : (
                    <UploadDropzone
                      icon={<CloudUploadIcon />}
                      accept={MEMORIAL_IMAGE_ACCEPT}
                      label={t('memorial.editor.sections.featuredPortrait')}
                      hint={t('memorial.editor.sections.featuredPortraitHint')}
                      onFiles={handlePortraitFiles}
                    />
                  )}
                </div>

                <div className={styles.gallerySection}>
                  <div className={styles.galleryHeader}>
                    <span className={styles.galleryLabel}>
                      {t('memorial.editor.sections.lifeGallery')}
                    </span>
                    {galleryItemCount > 0 && (
                      <span className={styles.galleryCount}>
                        {t('memorial.editor.sections.galleryCount', {
                          count: galleryItemCount,
                        })}
                      </span>
                    )}
                  </div>

                  {galleryItemCount > 0 ? (
                    <div className={styles.galleryGrid}>
                      {existingGalleryImages.map((image) => (
                        <article key={image.id} className={styles.galleryCard}>
                          <div className={styles.galleryPreview}>
                            {image.previewUrl ? (
                              <img
                                src={image.previewUrl}
                                alt={image.fileName}
                                className={styles.galleryPreviewImg}
                              />
                            ) : (
                              <div className={styles.galleryFallback}>
                                <GalleryIcon />
                              </div>
                            )}
                          </div>
                          <div className={styles.galleryCardBody}>
                            <strong>{image.fileName}</strong>
                            <span>
                              {buildFileMeta(image.fileSize, image.mimeType)}
                            </span>
                          </div>
                          <button
                            type="button"
                            className={styles.galleryCardRemove}
                            onClick={() => removeExistingGalleryImage(image.id)}
                          >
                            {t('memorial.editor.sections.removeGalleryImage')}
                          </button>
                        </article>
                      ))}

                      {newGalleryImages.map((image) => (
                        <article
                          key={image.localId}
                          className={`${styles.galleryCard} ${styles.galleryCardPending}`}
                        >
                          <div className={styles.galleryPreview}>
                            <img
                              src={image.previewUrl}
                              alt={image.file.name}
                              className={styles.galleryPreviewImg}
                            />
                          </div>
                          <div className={styles.galleryCardBody}>
                            <strong>{image.file.name}</strong>
                            <span>
                              {buildFileMeta(image.file.size, image.file.type)}
                            </span>
                          </div>
                          <button
                            type="button"
                            className={styles.galleryCardRemove}
                            onClick={() => removeNewGalleryImage(image.localId)}
                          >
                            {t('memorial.editor.sections.removeGalleryImage')}
                          </button>
                        </article>
                      ))}
                    </div>
                  ) : (
                    <div className={styles.galleryEmptyState}>
                      <span className={styles.galleryEmptyIcon}>
                        <GalleryIcon />
                      </span>
                      <strong>{t('memorial.editor.sections.galleryEmptyTitle')}</strong>
                      <p>{t('memorial.editor.sections.galleryEmptyText')}</p>
                    </div>
                  )}

                  <UploadDropzone
                    multiple
                    variant="compact"
                    icon={<GalleryIcon />}
                    accept={MEMORIAL_IMAGE_ACCEPT}
                    label={t('memorial.editor.sections.addGallery')}
                    hint={t('memorial.editor.sections.addGalleryHint')}
                    onFiles={handleGalleryFiles}
                  />
                </div>
              </div>
            </section>

            <EntryActions
              status={form.status === 'published' ? 'published' : 'draft'}
              publishLabel={t('memorial.editor.actions.publish')}
              saveDraftLabel={t('memorial.editor.actions.saveDraft')}
              deleteLabel={t('memorial.list.delete')}
              deleteConfirmTitle={t('memorial.list.deleteDialogTitle')}
              deleteConfirmBody={t('memorial.list.deleteDialogDescriptionPlain', {
                title: currentEntry?.fullName ?? '',
              })}
              deleteConfirmLabel={t('memorial.list.confirmDelete')}
              onPublish={() => void handleSave(true)}
              onSaveDraft={() => void handleSave(false)}
              onDelete={currentEntry ? () => void handleDelete() : undefined}
              isSubmitting={isSaving}
              isDeleting={isDeleting}
            />
          </div>

          <div className={styles.sideColumn}>
            <PublishingControls
              status={form.status === 'published' ? 'published' : 'draft'}
              publishOn={currentEntry?.publishedAt ? formatDate(currentEntry.publishedAt) : undefined}
            />

            <div className={styles.sideCard}>
              <span className={styles.sideCardLabel}>
                {t('memorial.editor.sections.lifeTimeline')}
              </span>
              <div className={styles.sideCardBody}>
                <div className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.dateOfBirth')}
                  </span>
                  <input
                    type="date"
                    value={form.dateOfBirth}
                    onChange={(event) =>
                      updateField('dateOfBirth', event.target.value)
                    }
                  />
                </div>
                <div className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.dateOfPassing')}
                  </span>
                  <input
                    type="date"
                    value={form.dateOfPassing}
                    onChange={(event) =>
                      updateField('dateOfPassing', event.target.value)
                    }
                  />
                </div>
              </div>
            </div>

            <div className={styles.sideCard}>
              <span className={styles.sideCardLabel}>
                {t('memorial.editor.sections.publishing')}
              </span>
              <div className={styles.sideCardBody}>
                <div className={styles.metaRow}>
                  <span className={styles.metaLabel}>
                    {t('resources.editor.meta.created')}
                  </span>
                  <strong>{formatDate(currentEntry?.createdAt)}</strong>
                </div>
                <div className={styles.metaRow}>
                  <span className={styles.metaLabel}>
                    {t('resources.editor.meta.updated')}
                  </span>
                  <strong>{formatDate(currentEntry?.updatedAt)}</strong>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </CmsAppShell>
  )
}

function buildFileMeta(fileSize: number, mimeType: string) {
  const typeLabel = mimeType.split('/')[1]?.toUpperCase() || 'IMAGE'

  if (fileSize < 1024) {
    return `${typeLabel} | ${fileSize} B`
  }
  if (fileSize < 1024 * 1024) {
    return `${typeLabel} | ${(fileSize / 1024).toFixed(1)} KB`
  }
  return `${typeLabel} | ${(fileSize / (1024 * 1024)).toFixed(1)} MB`
}

function GalleryIcon() {
  return (
    <svg viewBox="0 0 24 24" width="22" height="22" fill="none" aria-hidden="true">
      <rect x="3" y="3" width="18" height="18" rx="3" stroke="currentColor" strokeWidth="1.6" />
      <circle cx="8.5" cy="8.5" r="1.5" fill="currentColor" />
      <path d="M3 15l5-5 4 4 3-3 6 6" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}
