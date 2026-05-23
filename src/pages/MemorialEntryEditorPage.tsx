import { useEffect, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { EntryActions } from '../components/cms/EntryActions'
import { PublishingControls } from '../components/cms/PublishingControls'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { CloudUploadIcon } from '../components/icons'
import {
  MEMORIAL_CATEGORIES,
  MOCK_MEMORIAL_ENTRIES,
  type MemorialFormErrors,
  type MemorialFormState,
} from '../lib/memorialTypes'
import styles from '../styles/MemorialEntryEditorPage.module.css'

type MockGallery = {
  id: number
  name: string
  assetCount: number
  status: 'published' | 'draft'
}

const MOCK_GALLERIES: MockGallery[] = [
  { id: 1, name: '2025 CSAA Gathering', assetCount: 3, status: 'draft' },
  { id: 2, name: 'Annual Meeting', assetCount: 1, status: 'draft' },
  { id: 3, name: 'CSAA Reunions', assetCount: 15, status: 'published' },
  { id: 4, name: 'Partners and Donations', assetCount: 20, status: 'draft' },
  { id: 5, name: 'Taking the Children', assetCount: 6, status: 'published' },
]

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
    publishAt: '',
    dateOfBirth: '',
    dateOfPassing: '',
  }
}

function makeId() {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

export function MemorialEntryEditorPage({
  mode = 'create',
}: MemorialEntryEditorPageProps) {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()

  const isEditMode = mode === 'edit' && Boolean(id)

  const [isLoading, setIsLoading] = useState(isEditMode)
  const [isSaving, setIsSaving] = useState(false)
  const [form, setForm] = useState<MemorialFormState>(emptyForm())
  const [errors, setErrors] = useState<MemorialFormErrors>({})

  // Portrait
  const [portraitFile, setPortraitFile] = useState<File | null>(null)
  const [portraitPreviewUrl, setPortraitPreviewUrl] = useState<string | null>(null)

  // Gallery picker
  const [selectedGallery, setSelectedGallery] = useState<MockGallery | null>(null)
  const [galleryPickerOpen, setGalleryPickerOpen] = useState(false)

  useEffect(() => {
    if (!isEditMode) return

    const entry = MOCK_MEMORIAL_ENTRIES.find((e) => e.id === Number(id))
    if (entry) {
      setForm({
        fullName: entry.fullName,
        biography: entry.biography,
        category: entry.category,
        affiliation: entry.affiliation,
        status: entry.status,
        publishAt: '',
        dateOfBirth: entry.dateOfBirth ?? '',
        dateOfPassing: entry.dateOfPassing ?? '',
      })
    }
    setIsLoading(false)
  }, [isEditMode, id])

  function updateField<K extends keyof MemorialFormState>(
    key: K,
    value: MemorialFormState[K],
  ) {
    setForm((f) => ({ ...f, [key]: value }))
    if (errors[key as keyof MemorialFormErrors]) {
      setErrors((e) => ({ ...e, [key]: undefined }))
    }
  }

  function validate(): boolean {
    const next: MemorialFormErrors = {}
    if (!form.fullName.trim()) {
      next.fullName = t('memorial.editor.validation.fullNameRequired')
    }
    setErrors(next)
    return Object.keys(next).length === 0
  }

  async function handlePublish() {
    if (!validate()) return
    setIsSaving(true)
    await new Promise((r) => setTimeout(r, 600))
    updateField('status', 'published')
    setIsSaving(false)
    toast.success(
      isEditMode
        ? t('memorial.feedback.updated')
        : t('memorial.feedback.created'),
    )
    navigate('/memorial')
  }

  async function handleSaveDraft() {
    if (!form.fullName.trim()) {
      setErrors({ fullName: t('memorial.editor.validation.fullNameRequired') })
      return
    }
    setIsSaving(true)
    await new Promise((r) => setTimeout(r, 400))
    updateField('status', 'draft')
    setIsSaving(false)
    toast.success(t('memorial.feedback.updated'))
  }

  function handlePortraitFiles(files: File[]) {
    const file = files[0]
    if (!file) return
    if (portraitPreviewUrl) URL.revokeObjectURL(portraitPreviewUrl)
    setPortraitFile(file)
    setPortraitPreviewUrl(URL.createObjectURL(file))
  }

  function removePortrait() {
    if (portraitPreviewUrl) URL.revokeObjectURL(portraitPreviewUrl)
    setPortraitFile(null)
    setPortraitPreviewUrl(null)
  }

  const pageTitle = isEditMode
    ? t('memorial.editor.editTitle')
    : t('memorial.editor.createTitle')

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
            { label: isEditMode ? form.fullName || pageTitle : t('memorial.breadcrumb.createNew') },
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

        <div className={styles.layout}>
          {/* ── Main column ── */}
          <div className={styles.mainColumn}>

            {/* Full Name */}
            <div className={styles.card}>
              <div className={styles.field}>
                <span className={styles.fieldLabel}>
                  {t('memorial.editor.fields.fullName')}
                </span>
                <input
                  type="text"
                  value={form.fullName}
                  onChange={(e) => updateField('fullName', e.target.value)}
                  placeholder={t('memorial.editor.fields.fullNamePlaceholder')}
                />
                {errors.fullName && (
                  <p className={styles.fieldError}>{errors.fullName}</p>
                )}
              </div>
            </div>

            {/* Biography */}
            <div className={styles.biographyCard}>
              <p className={styles.biographyLabel}>
                {t('memorial.editor.sections.biography')}
              </p>
              <RichTextEditor
                value={form.biography}
                onChange={(html) => updateField('biography', html)}
                placeholder={t('memorial.editor.fields.biographyPlaceholder')}
                allowImages={false}
              />
            </div>

            {/* Media & Portrait */}
            <div className={styles.mediaCard}>
              <h2 className={styles.mediaCardTitle}>
                {t('memorial.editor.sections.media')}
              </h2>
              <div className={styles.mediaLayout}>
                {/* Featured Portrait */}
                <div className={styles.portraitZone}>
                  {portraitPreviewUrl ? (
                    <div className={styles.portraitPreview}>
                      <img
                        src={portraitPreviewUrl}
                        alt={form.fullName}
                        className={styles.portraitPreviewImg}
                      />
                      <button
                        type="button"
                        className={styles.portraitPreviewRemove}
                        onClick={removePortrait}
                        aria-label={t('memorial.editor.sections.removePortrait')}
                      >
                        ×
                      </button>
                    </div>
                  ) : (
                    <UploadDropzone
                      icon={<CloudUploadIcon />}
                      accept="image/png,image/jpeg"
                      label={t('memorial.editor.sections.featuredPortrait')}
                      hint={t('memorial.editor.sections.featuredPortraitHint')}
                      onFiles={handlePortraitFiles}
                    />
                  )}
                </div>

                {/* Life Gallery */}
                <div className={styles.gallerySection}>
                  <span className={styles.galleryLabel}>
                    {t('memorial.editor.sections.lifeGallery')}
                  </span>

                  {selectedGallery ? (
                    <div className={styles.selectedGalleryCard}>
                      <div className={styles.selectedGalleryIcon}>
                        <GalleryIcon />
                      </div>
                      <div className={styles.selectedGalleryInfo}>
                        <span className={styles.selectedGalleryName}>
                          {selectedGallery.name}
                        </span>
                        <span className={styles.selectedGalleryMeta}>
                          {selectedGallery.assetCount}{' '}
                          {selectedGallery.assetCount === 1 ? 'asset' : 'assets'}
                          {' · '}
                          <span className={selectedGallery.status === 'published'
                            ? styles.galleryStatusPublished
                            : styles.galleryStatusDraft}>
                            {selectedGallery.status === 'published' ? 'Published' : 'Draft'}
                          </span>
                        </span>
                      </div>
                      <div className={styles.selectedGalleryActions}>
                        <button
                          type="button"
                          className={styles.galleryChangeBtn}
                          onClick={() => setGalleryPickerOpen(true)}
                        >
                          {t('memorial.editor.sections.changeGallery')}
                        </button>
                        <button
                          type="button"
                          className={styles.galleryRemoveBtn}
                          onClick={() => setSelectedGallery(null)}
                        >
                          {t('memorial.editor.sections.removeGallery')}
                        </button>
                      </div>
                    </div>
                  ) : (
                    <button
                      type="button"
                      className={styles.galleryPickerTrigger}
                      onClick={() => setGalleryPickerOpen(true)}
                    >
                      <span className={styles.galleryPickerIcon}>
                        <GalleryIcon />
                      </span>
                      <span className={styles.galleryPickerLabel}>
                        {t('memorial.editor.sections.addGallery')}
                      </span>
                      <span className={styles.galleryPickerHint}>
                        {t('memorial.editor.sections.addGalleryHint')}
                      </span>
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* Action buttons — bottom of main column, matching other editor pages */}
            <EntryActions
              status={form.status === 'published' ? 'published' : 'draft'}
              publishLabel={t('memorial.editor.actions.publish')}
              saveDraftLabel={t('memorial.editor.actions.saveDraft')}
              onPublish={() => void handlePublish()}
              onSaveDraft={() => void handleSaveDraft()}
              isSubmitting={isSaving}
            />
          </div>

          {/* ── Sidebar ── */}
          <div className={styles.sideColumn}>

            {/* Publishing card */}
            <PublishingControls
              status={form.status === 'published' ? 'published' : 'draft'}
              publishOn={form.publishAt || undefined}
            />

            {/* Life Timeline card */}
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
                    onChange={(e) => updateField('dateOfBirth', e.target.value)}
                  />
                </div>
                <div className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('memorial.editor.fields.dateOfPassing')}
                  </span>
                  <input
                    type="date"
                    value={form.dateOfPassing}
                    onChange={(e) =>
                      updateField('dateOfPassing', e.target.value)
                    }
                  />
                </div>
              </div>
            </div>

          </div>
        </div>
      </div>

      {/* Gallery Picker Modal */}
      {galleryPickerOpen && (
        <div
          className={styles.pickerOverlay}
          onClick={() => setGalleryPickerOpen(false)}
        >
          <div
            className={styles.pickerModal}
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.pickerHeader}>
              <h2 className={styles.pickerTitle}>
                {t('memorial.editor.sections.galleryPickerTitle')}
              </h2>
              <button
                type="button"
                className={styles.pickerClose}
                onClick={() => setGalleryPickerOpen(false)}
                aria-label="Close"
              >
                ×
              </button>
            </div>
            <div className={styles.pickerList}>
              {MOCK_GALLERIES.map((gallery) => (
                <button
                  key={gallery.id}
                  type="button"
                  className={[
                    styles.pickerItem,
                    selectedGallery?.id === gallery.id
                      ? styles.pickerItemSelected
                      : '',
                  ]
                    .filter(Boolean)
                    .join(' ')}
                  onClick={() => {
                    setSelectedGallery(gallery)
                    setGalleryPickerOpen(false)
                  }}
                >
                  <span className={styles.pickerItemIcon}>
                    <GalleryIcon />
                  </span>
                  <span className={styles.pickerItemInfo}>
                    <span className={styles.pickerItemName}>{gallery.name}</span>
                    <span className={styles.pickerItemMeta}>
                      {gallery.assetCount}{' '}
                      {gallery.assetCount === 1 ? 'asset' : 'assets'}
                      {' · '}
                      <span className={gallery.status === 'published'
                        ? styles.galleryStatusPublished
                        : styles.galleryStatusDraft}>
                        {gallery.status === 'published' ? 'Published' : 'Draft'}
                      </span>
                    </span>
                  </span>
                  {selectedGallery?.id === gallery.id && (
                    <span className={styles.pickerItemCheck}>✓</span>
                  )}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </CmsAppShell>
  )
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
