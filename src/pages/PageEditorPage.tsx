import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import {
  CloudUploadIcon,
  EditIcon,
} from '../components/icons'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { pagesApi } from '../api/pagesApi'
import {
  buildPageFormStateFromDetail,
  buildSavePagePayload,
  createDefaultPageFormState,
  normalizePageSlugInput,
  validatePageForm,
  type PageFormErrors,
  type PageFormState,
} from '../lib/pagesForm'
import {
  clearCurrentPage,
  clearPageSaveState,
  createPage,
  fetchPageById,
  selectCurrentPage,
  selectPageDetail,
  selectPageSave,
  updatePage,
} from '../store/pagesSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import styles from '../styles/PageEditorPage.module.css'

type SubmitMode = 'draft' | 'publish'

type PageEditorPageProps = {
  mode?: 'edit' | 'view'
}

export function PageEditorPage({ mode = 'edit' }: PageEditorPageProps) {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const params = useParams()
  const { t } = useTranslation()
  const currentPage = useAppSelector(selectCurrentPage)
  const detailState = useAppSelector(selectPageDetail)
  const saveState = useAppSelector(selectPageSave)

  const parsedPageId = params.id ? Number.parseInt(params.id, 10) : null
  const isEditMode = parsedPageId !== null && Number.isFinite(parsedPageId)
  const isViewMode = mode === 'view' && isEditMode
  const isInvalidEditId = params.id !== undefined && !isEditMode

  const [form, setForm] = useState<PageFormState>(createDefaultPageFormState())
  const [errors, setErrors] = useState<PageFormErrors>({})
  const [slugTouched, setSlugTouched] = useState(false)
  const [heroImagePreviewUrl, setHeroImagePreviewUrl] = useState<string | null>(null)
  const [existingHeroObjectUrl, setExistingHeroObjectUrl] = useState<string | null>(null)

  useEffect(() => {
    if (isEditMode && parsedPageId) {
      void dispatch(fetchPageById(parsedPageId))
      return
    }

    dispatch(clearCurrentPage())
    setForm(createDefaultPageFormState())
    setSlugTouched(false)
  }, [dispatch, isEditMode, parsedPageId])

  useEffect(() => {
    if (isEditMode && currentPage) {
      setForm(buildPageFormStateFromDetail(currentPage))
      setErrors({})
      setSlugTouched(true)
    }
  }, [currentPage, isEditMode])

  useEffect(() => {
    return () => {
      dispatch(clearPageSaveState())
    }
  }, [dispatch])

  useEffect(() => {
    if (!form.heroImageFile) {
      setHeroImagePreviewUrl(null)
      return
    }

    const objectUrl = URL.createObjectURL(form.heroImageFile)
    setHeroImagePreviewUrl(objectUrl)

    return () => {
      URL.revokeObjectURL(objectUrl)
    }
  }, [form.heroImageFile])

  useEffect(() => {
    if (
      !form.heroImageEnabled ||
      form.removeHeroImage ||
      form.heroImageFile ||
      !form.existingHeroImageFetchUrl
    ) {
      setExistingHeroObjectUrl((current) => {
        if (current) {
          URL.revokeObjectURL(current)
        }
        return null
      })
      return
    }

    let cancelled = false
    let objectUrlToRevoke: string | null = null

    async function loadHeroImage() {
      try {
        const blob = await pagesApi.fetchPageHeroImageContent(form.existingHeroImageFetchUrl)
        const objectUrl = URL.createObjectURL(blob)
        objectUrlToRevoke = objectUrl

        if (cancelled) {
          URL.revokeObjectURL(objectUrl)
          return
        }

        setExistingHeroObjectUrl((current) => {
          if (current) {
            URL.revokeObjectURL(current)
          }
          return objectUrl
        })
      } catch {
        return
      }
    }

    void loadHeroImage()

    return () => {
      cancelled = true
      if (objectUrlToRevoke) {
        URL.revokeObjectURL(objectUrlToRevoke)
      }
    }
  }, [
    form.existingHeroImageFetchUrl,
    form.heroImageEnabled,
    form.heroImageFile,
    form.removeHeroImage,
  ])

  const isBusy = saveState.status === 'loading'
  const isInitialLoad = isEditMode && detailState.status === 'loading' && !currentPage
  const errorMessages = useMemo(
    () => Array.from(new Set(Object.values(errors).filter(Boolean))),
    [errors],
  )
  const currentHeroPreview =
    heroImagePreviewUrl ??
    (form.heroImageEnabled && !form.removeHeroImage ? existingHeroObjectUrl : null)

  function clearErrors(...keys: Array<keyof PageFormState>) {
    setErrors((current) => {
      const next = { ...current }
      for (const key of keys) {
        delete next[key]
      }
      return next
    })
  }

  function updateField<K extends keyof PageFormState>(
    key: K,
    value: PageFormState[K],
  ) {
    setForm((current) => ({
      ...current,
      [key]: value,
    }))
    clearErrors(key)
  }

  function handlePageTitleChange(value: string) {
    setForm((current) => ({
      ...current,
      pageTitle: value,
      urlSlug: slugTouched ? current.urlSlug : normalizePageSlugInput(value),
    }))
    clearErrors('pageTitle', 'urlSlug')
  }

  function handleSlugChange(value: string) {
    setSlugTouched(true)
    updateField('urlSlug', normalizePageSlugInput(value))
  }

  function handleStatusToggle(nextPublished: boolean) {
    updateField('status', nextPublished ? 'published' : 'draft')
  }

  function handleHeroToggle(enabled: boolean) {
    setForm((current) => ({
      ...current,
      heroImageEnabled: enabled,
      heroImageFile: enabled ? current.heroImageFile : null,
      removeHeroImage: enabled ? false : true,
    }))
    clearErrors('heroImageFile')
  }

  function handleHeroFiles(files: File[]) {
    const file = files[0]
    if (!file) {
      return
    }

    setForm((current) => ({
      ...current,
      heroImageEnabled: true,
      heroImageFile: file,
      removeHeroImage: false,
    }))
    clearErrors('heroImageFile')
  }

  function removeHeroImage() {
    setForm((current) => ({
      ...current,
      heroImageFile: null,
      removeHeroImage: true,
    }))
  }

  async function submitForm(submitMode: SubmitMode) {
    const submissionState: PageFormState = {
      ...form,
      status: submitMode === 'publish' ? 'published' : 'draft',
    }

    const validationErrors = validatePageForm(submissionState, t)
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('pages.feedback.validation'))
      return
    }

    setErrors({})

    try {
      const payload = await buildSavePagePayload(submissionState)
      const result =
        isEditMode && parsedPageId
          ? await dispatch(updatePage({ id: parsedPageId, payload })).unwrap()
          : await dispatch(createPage(payload)).unwrap()

      if (isEditMode && parsedPageId) {
        await dispatch(fetchPageById(parsedPageId)).unwrap()
        toast.success(
          submitMode === 'publish'
            ? t('pages.feedback.published')
            : t('pages.feedback.draftSaved'),
        )
      } else {
        toast.success(t('pages.feedback.created'))
        navigate(`/pages/${result.page.id}/edit`, { replace: true })
      }
    } catch {
      return
    }
  }

  if (isInvalidEditId) {
    return (
      <CmsAppShell activeKey="pages">
        <section className={styles.errorPanel}>
          <h1>{t('pages.editor.invalidIdTitle')}</h1>
          <p>{t('pages.editor.invalidIdText')}</p>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => navigate('/pages')}
          >
            {t('pages.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  if (isInitialLoad) {
    return (
      <CmsAppShell activeKey="pages">
        <div className={styles.loaderWrap}>
          <Loader label={t('pages.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (isEditMode && detailState.error && !currentPage) {
    return (
      <CmsAppShell activeKey="pages">
        <section className={styles.errorPanel}>
          <h1>{t('pages.editor.loadErrorTitle')}</h1>
          <p>{detailState.error}</p>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => navigate('/pages')}
          >
            {t('pages.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="pages">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('pages.breadcrumb.pages'), to: '/pages' },
            {
              label: isViewMode
                ? t('pages.editor.breadcrumbView')
                : isEditMode
                  ? t('pages.editor.breadcrumbEdit')
                  : t('pages.editor.breadcrumbCreate'),
            },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>
              {isViewMode
                ? t('pages.editor.titleView')
                : isEditMode
                  ? t('pages.editor.titleEdit')
                  : t('pages.editor.titleCreate')}
            </h1>
            <p className={styles.subtitle}>{t('pages.editor.subtitle')}</p>
          </div>

          <div className={styles.headerActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => navigate('/pages')}
            >
              {t('pages.editor.backToList')}
            </button>
            {isViewMode && parsedPageId && (
              <button
                type="button"
                className={styles.primaryButton}
                onClick={() => navigate(`/pages/${parsedPageId}/edit`)}
              >
                <EditIcon size={14} />
                {t('pages.editor.editPage')}
              </button>
            )}
          </div>
        </header>

        {!isViewMode && errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{t('pages.feedback.validation')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        {!isViewMode && saveState.error && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{saveState.error}</p>
          </div>
        )}

        <fieldset className={styles.readOnlyFieldset} disabled={isViewMode || isBusy}>
          <section className={styles.card}>
            <div className={styles.cardHeader}>
              <h2>{t('pages.sections.pageSettings')}</h2>
              <div className={styles.statusSwitch}>
                <span
                  className={form.status === 'draft' ? styles.statusLabelActive : styles.statusLabel}
                >
                  {t('pages.status.draft')}
                </span>
                <button
                  type="button"
                  className={styles.switch}
                  role="switch"
                  aria-checked={form.status === 'published'}
                  onClick={() => handleStatusToggle(form.status !== 'published')}
                >
                  <span
                    className={[
                      styles.switchThumb,
                      form.status === 'published' ? styles.switchThumbOn : '',
                    ].join(' ')}
                  />
                </button>
                <span
                  className={
                    form.status === 'published' ? styles.statusLabelActive : styles.statusLabel
                  }
                >
                  {t('pages.status.live')}
                </span>
              </div>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.fields.pageTitle')}</span>
                <input
                  type="text"
                  value={form.pageTitle}
                  placeholder={t('pages.fields.pageTitlePlaceholder')}
                  onChange={(event) => handlePageTitleChange(event.target.value)}
                />
                <FieldError message={errors.pageTitle} />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.urlSlug')}</span>
                <div className={styles.slugField}>
                  <span className={styles.slugPrefix}>/</span>
                  <input
                    type="text"
                    value={form.urlSlug}
                    placeholder={t('pages.fields.urlSlugPlaceholder')}
                    onChange={(event) => handleSlugChange(event.target.value)}
                  />
                </div>
                <FieldError message={errors.urlSlug} />
              </label>
            </div>

            <div className={styles.heroSection}>
              <label className={styles.toggleRow}>
                <div>
                  <span>{t('pages.fields.enableHeroImage')}</span>
                  <p>{t('pages.fields.enableHeroImageHint')}</p>
                </div>
                <input
                  type="checkbox"
                  checked={form.heroImageEnabled}
                  onChange={(event) => handleHeroToggle(event.target.checked)}
                />
              </label>

              {form.heroImageEnabled && (
                <div className={styles.heroPanel}>
                  {currentHeroPreview ? (
                    <div className={styles.heroPreviewCard}>
                      <img
                        src={currentHeroPreview}
                        alt={form.pageTitle || t('pages.fields.heroImageAlt')}
                        className={styles.heroPreview}
                      />
                      {!isViewMode && (
                        <div className={styles.heroActions}>
                          <label className={styles.secondaryButton}>
                            <input
                              type="file"
                              accept="image/png,image/jpeg,image/webp"
                              hidden
                              onChange={(event) => {
                                const file = event.target.files?.[0]
                                if (file) {
                                  handleHeroFiles([file])
                                }
                                event.target.value = ''
                              }}
                            />
                            {t('pages.editor.replaceHeroImage')}
                          </label>
                          <button
                            type="button"
                            className={styles.dangerButton}
                            onClick={removeHeroImage}
                          >
                            {t('pages.editor.removeHeroImage')}
                          </button>
                        </div>
                      )}
                    </div>
                  ) : (
                    <UploadDropzone
                      icon={<CloudUploadIcon size={24} />}
                      accept="image/png,image/jpeg,image/webp"
                      label={t('pages.fields.heroImageLabel')}
                      hint={t('pages.fields.heroImageHint')}
                      onFiles={handleHeroFiles}
                    />
                  )}
                </div>
              )}
            </div>
          </section>

          <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>{t('pages.sections.metadata')}</h2>
              <p>{t('pages.sections.metadataHint')}</p>
            </div>

            <div className={styles.fieldStack}>
              <label className={styles.field}>
                <span>{t('pages.fields.seoPageTitle')}</span>
                <input
                  type="text"
                  value={form.seoPageTitle}
                  placeholder={t('pages.fields.seoPageTitlePlaceholder')}
                  onChange={(event) => updateField('seoPageTitle', event.target.value)}
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.seoPageDescription')}</span>
                <textarea
                  rows={5}
                  value={form.seoPageDescription}
                  placeholder={t('pages.fields.seoPageDescriptionPlaceholder')}
                  onChange={(event) => updateField('seoPageDescription', event.target.value)}
                />
              </label>
            </div>
          </section>
        </fieldset>

        {!isViewMode && (
          <section className={styles.actionsCard}>
            <div className={styles.actionRow}>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={isBusy}
                onClick={() => submitForm('draft')}
              >
                {t('pages.editor.saveDraft')}
              </button>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={isBusy}
                onClick={() => submitForm('publish')}
              >
                {isBusy ? t('pages.common.loading') : t('pages.editor.publish')}
              </button>
            </div>
          </section>
        )}
      </div>
    </CmsAppShell>
  )
}

function FieldError({ message }: { message?: string }) {
  if (!message) {
    return null
  }

  return <span className={styles.fieldError}>{message}</span>
}
