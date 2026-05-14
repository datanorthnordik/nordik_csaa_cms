import { useEffect, useMemo, useState, type DragEvent, type ReactNode } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import { mediaApi } from '../api/mediaApi'
import {
  isModulePage,
  pagesApi,
  resolvePageParentId,
  resolvePageParentSlug,
  type PageParentOption,
  type PageSectionType,
} from '../api/pagesApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import {
  AddIcon,
  CloudUploadIcon,
  DeleteIcon,
  DownloadIcon,
  EditIcon,
  GalleryIcon,
  PagesIcon,
} from '../components/icons'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import {
  buildFullPageUrlSlug,
  buildPageFormStateFromDetail,
  buildSavePageRequest,
  createDefaultDocumentState,
  createDefaultPageFormState,
  createDefaultSectionState,
  getDisallowedParentPageIds,
  normalizePageSlugInput,
  reorderPageSections,
  validatePageForm,
  type PageFormErrors,
  type PageFormState,
  type PageSectionDocumentState,
  type PageSectionState,
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

type PageOptionsStatus = 'idle' | 'loading' | 'succeeded' | 'failed'
type GalleryOptionsStatus = 'idle' | 'loading' | 'succeeded' | 'failed'

type PageGalleryOption = {
  id: number
  name: string
}

type ModuleTypeOption = {
  type: PageSectionType
  label: string
  description: string
  icon: ReactNode
}

const DOCUMENT_UPLOAD_ACCEPT = [
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
  const isModuleManaged = currentPage ? isModulePage(currentPage) : false
  const isReadOnlyMode = isViewMode || isModuleManaged

  const [form, setForm] = useState<PageFormState>(createDefaultPageFormState())
  const [errors, setErrors] = useState<PageFormErrors>({})
  const [slugTouched, setSlugTouched] = useState(false)
  const [heroImagePreviewUrl, setHeroImagePreviewUrl] = useState<string | null>(null)
  const [existingHeroObjectUrl, setExistingHeroObjectUrl] = useState<string | null>(null)
  const [parentPageOptions, setParentPageOptions] = useState<PageParentOption[]>([])
  const [parentPageOptionsStatus, setParentPageOptionsStatus] =
    useState<PageOptionsStatus>('idle')
  const [parentPageOptionsError, setParentPageOptionsError] = useState<string | null>(null)
  const [galleryOptions, setGalleryOptions] = useState<PageGalleryOption[]>([])
  const [galleryOptionsStatus, setGalleryOptionsStatus] =
    useState<GalleryOptionsStatus>('idle')
  const [galleryOptionsError, setGalleryOptionsError] = useState<string | null>(null)
  const [initializedPageId, setInitializedPageId] = useState<number | null>(null)
  const [modulePickerOpen, setModulePickerOpen] = useState(false)
  const [draggedSectionClientId, setDraggedSectionClientId] = useState<string | null>(null)
  const [dragOverSectionClientId, setDragOverSectionClientId] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false

    async function loadParentPageOptions() {
      setParentPageOptionsStatus('loading')
      setParentPageOptionsError(null)

      try {
        const options = await pagesApi.listPageParentOptions()
        if (cancelled) {
          return
        }

        setParentPageOptions(options)
        setParentPageOptionsStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setParentPageOptions([])
        setParentPageOptionsStatus('failed')
        setParentPageOptionsError(getApiErrorMessage(error))
      }
    }

    async function loadGalleryOptions() {
      setGalleryOptionsStatus('loading')
      setGalleryOptionsError(null)

      try {
        const options = await mediaApi.listGalleries()
        if (cancelled) {
          return
        }

        setGalleryOptions(
          options.map((gallery) => ({
            id: gallery.id,
            name: gallery.name,
          })),
        )
        setGalleryOptionsStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setGalleryOptions([])
        setGalleryOptionsStatus('failed')
        setGalleryOptionsError(getApiErrorMessage(error))
      }
    }

    void loadParentPageOptions()
    void loadGalleryOptions()

    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (isEditMode && parsedPageId) {
      setInitializedPageId(null)
      void dispatch(fetchPageById(parsedPageId))
      return
    }

    dispatch(clearCurrentPage())
    setForm(createDefaultPageFormState())
    setSlugTouched(false)
    setInitializedPageId(null)
  }, [dispatch, isEditMode, parsedPageId])

  const parentPageOptionsById = useMemo(
    () => new Map(parentPageOptions.map((page) => [page.id, page])),
    [parentPageOptions],
  )
  const currentPageParentId = currentPage ? resolvePageParentId(currentPage) : null
  const currentPageParentSlug =
    (currentPageParentId ? parentPageOptionsById.get(currentPageParentId)?.url_slug : '') ||
    (currentPage ? resolvePageParentSlug(currentPage) : '')
  const isWaitingForParentContext = Boolean(
    isEditMode &&
      currentPage &&
      currentPageParentId &&
      !currentPageParentSlug &&
      parentPageOptionsStatus === 'loading',
  )

  useEffect(() => {
    if (!isEditMode || !currentPage || initializedPageId === currentPage.id) {
      return
    }

    if (isWaitingForParentContext) {
      return
    }

    setForm(
      buildPageFormStateFromDetail(currentPage, {
        parentPageSlug: currentPageParentSlug,
      }),
    )
    setErrors({})
    setSlugTouched(true)
    setInitializedPageId(currentPage.id)
  }, [
    currentPage,
    currentPageParentId,
    currentPageParentSlug,
    initializedPageId,
    isEditMode,
    isWaitingForParentContext,
  ])

  const selectedParentPageId = form.parentPageId ? Number.parseInt(form.parentPageId, 10) : null
  const selectedParentPageSlug =
    (selectedParentPageId ? parentPageOptionsById.get(selectedParentPageId)?.url_slug : '') ||
    (selectedParentPageId === currentPageParentId ? currentPageParentSlug : '')
  const effectiveUrlSlug = buildFullPageUrlSlug(form.urlSlug, selectedParentPageSlug)
  const slugPrefix = selectedParentPageSlug
    ? `${selectedParentPageSlug.replace(/\/+$/, '')}/`
    : '/'
  const disallowedParentPageIds = useMemo(
    () =>
      isEditMode && parsedPageId
        ? getDisallowedParentPageIds(parsedPageId, parentPageOptions)
        : new Set<number>(),
    [isEditMode, parsedPageId, parentPageOptions],
  )
  const availableParentPageOptions = useMemo(
    () => parentPageOptions.filter((page) => page.id !== parsedPageId),
    [parentPageOptions, parsedPageId],
  )

  const moduleTypeOptions = useMemo<ModuleTypeOption[]>(
    () => [
      {
        type: 'typography',
        label: t('pages.modules.types.typography'),
        description: t('pages.modules.descriptions.typography'),
        icon: <TypographyModuleIcon />,
      },
      {
        type: 'gallery',
        label: t('pages.modules.types.gallery'),
        description: t('pages.modules.descriptions.gallery'),
        icon: <GalleryModuleIcon />,
      },
      {
        type: 'document',
        label: t('pages.modules.types.document'),
        description: t('pages.modules.descriptions.document'),
        icon: <DocumentModuleIcon />,
      },
      {
        type: 'quote',
        label: t('pages.modules.types.quote'),
        description: t('pages.modules.descriptions.quote'),
        icon: <QuoteModuleIcon />,
      },
      {
        type: 'cta_banner',
        label: t('pages.modules.types.cta_banner'),
        description: t('pages.modules.descriptions.cta_banner'),
        icon: <CtaModuleIcon />,
      },
      {
        type: 'header',
        label: t('pages.modules.types.header'),
        description: t('pages.modules.descriptions.header'),
        icon: <HeaderModuleIcon />,
      },
    ],
    [t],
  )

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
  const isInitialLoad =
    (isEditMode && detailState.status === 'loading' && !currentPage) || isWaitingForParentContext
  const errorMessages = useMemo(
    () => Array.from(new Set(Object.values(errors).filter(Boolean))),
    [errors],
  )
  const currentHeroPreview =
    heroImagePreviewUrl ??
    (form.heroImageEnabled && !form.removeHeroImage ? existingHeroObjectUrl : null)

  function clearErrors(...keys: Array<string>) {
    setErrors((current) => {
      const next = { ...current }
      for (const key of keys) {
        delete next[key as keyof PageFormErrors]
      }
      return next
    })
  }

  function updateField<K extends keyof PageFormState>(key: K, value: PageFormState[K]) {
    setForm((current) => ({
      ...current,
      [key]: value,
    }))
    clearErrors(key)
  }

  function updateSection(
    sectionClientId: string,
    updater: (section: PageSectionState) => PageSectionState,
  ) {
    setForm((current) => ({
      ...current,
      sections: current.sections.map((section) =>
        section.clientId === sectionClientId ? updater(section) : section,
      ),
    }))
    clearErrors('sections')
  }

  function updateDocument(
    sectionClientId: string,
    documentClientId: string,
    updater: (document: PageSectionDocumentState) => PageSectionDocumentState,
  ) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: section.documents.items.map((document) =>
          document.clientId === documentClientId ? updater(document) : document,
        ),
      },
    }))
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

  function handleParentPageChange(value: string) {
    updateField('parentPageId', value)
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

  function handleAddSection(sectionType: PageSectionType) {
    setForm((current) => {
      const nextSection = createDefaultSectionState(sectionType)
      if (sectionType === 'header' && current.sections.length > 0) {
        nextSection.header.hierarchy = 'h2_section'
      }
      return {
        ...current,
        sections: [...current.sections, nextSection],
      }
    })
    setModulePickerOpen(false)
    clearErrors('sections')
  }

  function handleRemoveSection(sectionClientId: string) {
    setForm((current) => ({
      ...current,
      sections: current.sections.filter((section) => section.clientId !== sectionClientId),
    }))
    clearErrors('sections')
  }

  function handleMoveSection(sectionClientId: string, direction: -1 | 1) {
    setForm((current) => ({
      ...current,
      sections: moveItemByClientId(current.sections, sectionClientId, direction),
    }))
    clearErrors('sections')
  }

  function handleSectionDragStart(sectionClientId: string) {
    if (isReadOnlyMode || isBusy) {
      return
    }
    setDraggedSectionClientId(sectionClientId)
  }

  function handleSectionDragOver(event: DragEvent<HTMLElement>, sectionClientId: string) {
    if (
      isReadOnlyMode ||
      isBusy ||
      !draggedSectionClientId ||
      draggedSectionClientId === sectionClientId
    ) {
      return
    }

    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
    setDragOverSectionClientId(sectionClientId)
  }

  function handleSectionDrop(sectionClientId: string) {
    if (
      isReadOnlyMode ||
      isBusy ||
      !draggedSectionClientId ||
      draggedSectionClientId === sectionClientId
    ) {
      return
    }

    setForm((current) => ({
      ...current,
      sections: reorderPageSections(current.sections, draggedSectionClientId, sectionClientId),
    }))
    setDraggedSectionClientId(null)
    setDragOverSectionClientId(null)
    clearErrors('sections')
  }

  function handleSectionDragEnd() {
    setDraggedSectionClientId(null)
    setDragOverSectionClientId(null)
  }

  function handleDocumentFiles(sectionClientId: string, files: File[]) {
    if (!files.length) {
      return
    }

    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: [...section.documents.items, ...files.map(createDefaultDocumentState)],
      },
    }))
  }

  function handleReplaceDocumentFile(
    sectionClientId: string,
    documentClientId: string,
    file: File,
  ) {
    updateDocument(sectionClientId, documentClientId, (document) => ({
      ...document,
      file,
      originalFileName: file.name,
      displayName: document.displayName.trim() || stripFileExtension(file.name),
      existingMimeType: file.type || document.existingMimeType,
      existingFileSize: file.size,
    }))
  }

  function handleRemoveDocument(sectionClientId: string, documentClientId: string) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: section.documents.items.filter((document) => document.clientId !== documentClientId),
      },
    }))
  }

  function handleMoveDocument(
    sectionClientId: string,
    documentClientId: string,
    direction: -1 | 1,
  ) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: moveItemByClientId(section.documents.items, documentClientId, direction),
      },
    }))
  }

  async function submitForm(submitMode: SubmitMode) {
    const submissionState: PageFormState = {
      ...form,
      status: submitMode === 'publish' ? 'published' : 'draft',
    }

    const validationErrors = validatePageForm(submissionState, t, {
      currentPageId: isEditMode ? parsedPageId : null,
      pageOptions: parentPageOptions,
    })
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('pages.feedback.validation'))
      return
    }

    setErrors({})

    try {
      const payload = buildSavePageRequest(submissionState, selectedParentPageSlug)
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

  function renderSectionBody(section: PageSectionState) {
    switch (section.sectionType) {
      case 'header':
        return (
          <div className={[styles.fieldStack, styles.documentSection].join(' ')}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.mainHeaderText')}</span>
                <input
                  type="text"
                  value={section.header.mainHeaderText}
                  placeholder={t('pages.modules.placeholders.mainHeaderText')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        mainHeaderText: event.target.value,
                      },
                    }))
                  }
                />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.subHeaderText')}</span>
                <input
                  type="text"
                  value={section.header.subHeaderText}
                  placeholder={t('pages.modules.placeholders.subHeaderText')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        subHeaderText: event.target.value,
                      },
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.hierarchy')}</span>
                <div className={styles.segmentedControl}>
                  <SegmentedButton
                    active={section.header.hierarchy === 'h1_hero'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        header: {
                          ...current.header,
                          hierarchy: 'h1_hero',
                        },
                      }))
                    }
                    label={t('pages.modules.options.h1Hero')}
                  />
                  <SegmentedButton
                    active={section.header.hierarchy === 'h2_section'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        header: {
                          ...current.header,
                          hierarchy: 'h2_section',
                        },
                      }))
                    }
                    label={t('pages.modules.options.h2Section')}
                  />
                </div>
              </div>
            </div>
          </div>
        )

      case 'typography':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.textAlign')}</span>
                <div className={styles.segmentedControl}>
                  <SegmentedButton
                    active={section.typography.textAlign === 'left'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'left',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignLeft')}
                  />
                  <SegmentedButton
                    active={section.typography.textAlign === 'center'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'center',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignCenter')}
                  />
                  <SegmentedButton
                    active={section.typography.textAlign === 'right'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'right',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignRight')}
                  />
                </div>
              </div>
            </div>

            <div className={styles.field}>
              <span>{t('pages.modules.fields.htmlContent')}</span>
              <RichTextEditor
                value={section.typography.htmlContent}
                onChange={(html) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    typography: {
                      ...current.typography,
                      htmlContent: html,
                    },
                  }))
                }
                placeholder={t('pages.modules.placeholders.htmlContent')}
                disabled={isReadOnlyMode || isBusy}
              />
            </div>
          </div>
        )

      case 'gallery':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.gallery')}</span>
                <select
                  value={section.gallery.galleryId}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        galleryId: event.target.value,
                      },
                    }))
                  }
                  disabled={galleryOptionsStatus === 'loading'}
                >
                  <option value="">{t('pages.modules.placeholders.gallery')}</option>
                  {galleryOptions.map((gallery) => (
                    <option key={gallery.id} value={gallery.id}>
                      {gallery.name}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>
                  {galleryOptionsError || t('pages.modules.hints.gallery')}
                </p>
              </label>
            </div>

            <div className={styles.field}>
              <span>{t('pages.modules.fields.viewMode')}</span>
              <div className={styles.segmentedControl}>
                <SegmentedButton
                  active={section.gallery.viewMode === 'grid'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        viewMode: 'grid',
                      },
                    }))
                  }
                  label={t('pages.modules.options.grid')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'carousel'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        viewMode: 'carousel',
                      },
                    }))
                  }
                  label={t('pages.modules.options.carousel')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'masonry'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        viewMode: 'masonry',
                      },
                    }))
                  }
                  label={t('pages.modules.options.masonry')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'focus'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        viewMode: 'focus',
                      },
                    }))
                  }
                  label={t('pages.modules.options.focus')}
                />
              </div>
            </div>
          </div>
        )

      case 'document':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.documents')}</span>
                <p className={styles.fieldHint}>{t('pages.modules.hints.documents')}</p>
              </div>
            </div>

            <UploadDropzone
              multiple
              accept={DOCUMENT_UPLOAD_ACCEPT}
              variant="compact"
              className={styles.documentDropzone}
              disabled={isReadOnlyMode || isBusy}
              icon={<CloudUploadIcon size={20} />}
              label={t('pages.modules.document.dropLabel')}
              hint={t('pages.modules.document.dropHint')}
              onFiles={(files) => handleDocumentFiles(section.clientId, files)}
            />

            {section.documents.items.length > 0 && (
              <div className={styles.documentList}>
                {section.documents.items.map((document, documentIndex) => {
                  const existingFileName =
                    document.originalFileName ||
                    document.file?.name ||
                    t('pages.modules.document.untitled')
                  const activeFileSize = document.file?.size ?? document.existingFileSize

                  return (
                    <article key={document.clientId} className={styles.documentItem}>
                      <div className={styles.documentHeader}>
                        <div className={styles.documentHeaderContent}>
                          <p className={styles.documentLabel}>
                            {t('pages.modules.document.itemLabel', {
                              index: documentIndex + 1,
                            })}
                          </p>
                          <h4 className={styles.documentName}>
                            {document.displayName || stripFileExtension(existingFileName)}
                          </h4>
                          <p className={styles.documentMeta}>
                            {document.file
                              ? t('pages.modules.document.pendingUpload', {
                                  name: document.file.name,
                                })
                              : existingFileName}
                            {activeFileSize ? ` - ${formatFileSize(activeFileSize)}` : ''}
                          </p>
                        </div>

                        {!isReadOnlyMode && (
                          <div className={styles.documentActionRow}>
                            <label className={styles.secondaryButton}>
                              <input
                                type="file"
                                hidden
                                accept={DOCUMENT_UPLOAD_ACCEPT}
                                onChange={(event) => {
                                  const file = event.target.files?.[0]
                                  if (file) {
                                    handleReplaceDocumentFile(
                                      section.clientId,
                                      document.clientId,
                                      file,
                                    )
                                  }
                                  event.target.value = ''
                                }}
                              />
                              <EditIcon size={14} />
                              {t('pages.modules.actions.replaceFile')}
                            </label>
                            <button
                              type="button"
                              className={styles.iconButton}
                              disabled={documentIndex === 0}
                              aria-label={t('pages.modules.actions.moveUp')}
                              onClick={() =>
                                handleMoveDocument(section.clientId, document.clientId, -1)
                              }
                            >
                              <ArrowUpIcon />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              disabled={documentIndex === section.documents.items.length - 1}
                              aria-label={t('pages.modules.actions.moveDown')}
                              onClick={() =>
                                handleMoveDocument(section.clientId, document.clientId, 1)
                              }
                            >
                              <ArrowDownIcon />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label={t('pages.modules.actions.removeDocument')}
                              onClick={() =>
                                handleRemoveDocument(section.clientId, document.clientId)
                              }
                            >
                              <DeleteIcon size={16} />
                            </button>
                          </div>
                        )}
                      </div>

                      <div className={styles.fieldGrid}>
                        <label className={styles.field}>
                          <span>{t('pages.modules.fields.documentDisplayName')}</span>
                          <input
                            type="text"
                            value={document.displayName}
                            placeholder={t('pages.modules.placeholders.documentDisplayName')}
                            onChange={(event) =>
                              updateDocument(section.clientId, document.clientId, (current) => ({
                                ...current,
                                displayName: event.target.value,
                              }))
                            }
                          />
                        </label>

                        <label className={styles.field}>
                          <span>{t('pages.modules.fields.documentDescription')}</span>
                          <input
                            type="text"
                            value={document.description}
                            placeholder={t('pages.modules.placeholders.documentDescription')}
                            onChange={(event) =>
                              updateDocument(section.clientId, document.clientId, (current) => ({
                                ...current,
                                description: event.target.value,
                              }))
                            }
                          />
                        </label>
                      </div>

                      {document.existingFetchUrl && !document.file && (
                        <div className={styles.documentFooter}>
                          <a
                            href={document.existingFetchUrl}
                            target="_blank"
                            rel="noreferrer"
                            className={styles.secondaryButton}
                          >
                            <DownloadIcon size={14} />
                            {t('pages.modules.actions.openDocument')}
                          </a>
                        </div>
                      )}
                    </article>
                  )
                })}
              </div>
            )}
          </div>
        )

      case 'quote':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.attribution')}</span>
                <input
                  type="text"
                  value={section.quote.attribution}
                  placeholder={t('pages.modules.placeholders.attribution')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      quote: {
                        ...current.quote,
                        attribution: event.target.value,
                      },
                    }))
                  }
                />
              </label>
            </div>

            <label className={styles.field}>
              <span>{t('pages.modules.fields.quoteContent')}</span>
              <textarea
                rows={5}
                value={section.quote.quoteContent}
                placeholder={t('pages.modules.placeholders.quoteContent')}
                onChange={(event) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    quote: {
                      ...current.quote,
                      quoteContent: event.target.value,
                    },
                  }))
                }
              />
            </label>
          </div>
        )

      case 'cta_banner':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.bannerHeading')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.bannerHeading}
                  placeholder={t('pages.modules.placeholders.bannerHeading')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        bannerHeading: event.target.value,
                      },
                    }))
                  }
                />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.bannerMessage')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.bannerMessage}
                  placeholder={t('pages.modules.placeholders.bannerMessage')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        bannerMessage: event.target.value,
                      },
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.buttonText')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.buttonText}
                  placeholder={t('pages.modules.placeholders.buttonText')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        buttonText: event.target.value,
                      },
                    }))
                  }
                />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.buttonUrl')}</span>
                <input
                  type="url"
                  value={section.ctaBanner.buttonUrl}
                  placeholder={t('pages.modules.placeholders.buttonUrl')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        buttonUrl: event.target.value,
                      },
                    }))
                  }
                />
              </label>

              <label className={styles.toggleRow}>
                <div>
                  <span>{t('pages.modules.fields.openInNewTab')}</span>
                  <p>{t('pages.modules.hints.openInNewTab')}</p>
                </div>
                <input
                  type="checkbox"
                  checked={section.ctaBanner.openInNewTab}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        openInNewTab: event.target.checked,
                      },
                    }))
                  }
                />
              </label>
            </div>
          </div>
        )
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
              label: isReadOnlyMode
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
              {isReadOnlyMode
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
            {isViewMode && parsedPageId && !isModuleManaged && (
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

        {isModuleManaged && (
          <div className={styles.infoSummary} role="status">
            <p className={styles.infoSummaryTitle}>{t('pages.editor.moduleManagedTitle')}</p>
            <p className={styles.infoSummaryText}>{t('pages.editor.moduleManagedNotice')}</p>
          </div>
        )}

        {!isReadOnlyMode && errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{t('pages.feedback.validation')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        {!isReadOnlyMode && saveState.error && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{saveState.error}</p>
          </div>
        )}

        <fieldset className={styles.readOnlyFieldset} disabled={isReadOnlyMode || isBusy}>
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
                <span>{t('pages.fields.parentPage')}</span>
                <select
                  value={form.parentPageId}
                  onChange={(event) => handleParentPageChange(event.target.value)}
                  disabled={parentPageOptionsStatus === 'loading'}
                >
                  <option value="">{t('pages.fields.noParentPage')}</option>
                  {availableParentPageOptions.map((page) => (
                    <option
                      key={page.id}
                      value={page.id}
                      disabled={disallowedParentPageIds.has(page.id)}
                    >
                      {`${page.page_title} (${page.url_slug})`}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>{t('pages.fields.parentPageHint')}</p>
                <FieldError message={errors.parentPageId ?? parentPageOptionsError ?? undefined} />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.urlSlug')}</span>
                <div className={styles.slugField}>
                  <span className={styles.slugPrefix}>{slugPrefix}</span>
                  <input
                    type="text"
                    value={form.urlSlug}
                    placeholder={t('pages.fields.urlSlugPlaceholder')}
                    onChange={(event) => handleSlugChange(event.target.value)}
                  />
                </div>
                <p className={styles.fieldHint}>
                  {t('pages.fields.fullUrlSlugPreview', {
                    slug: effectiveUrlSlug || '/',
                  })}
                </p>
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
                      {!isReadOnlyMode && (
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
                      disabled={isReadOnlyMode || isBusy}
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

          <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>{t('pages.sections.contentModules')}</h2>
              <p>{t('pages.sections.contentModulesHint')}</p>
            </div>

            {form.sections.length === 0 ? (
              <div className={styles.emptyState}>
                <div className={styles.emptyStateIcon}>
                  <PagesIcon size={18} />
                </div>
                <div>
                  <p className={styles.emptyStateTitle}>
                    {t('pages.sections.emptyModulesTitle')}
                  </p>
                  <p className={styles.emptyStateText}>
                    {t('pages.sections.emptyModulesText')}
                  </p>
                </div>
              </div>
            ) : (
              <div className={styles.moduleList}>
                {form.sections.map((section, sectionIndex) => {
                  const moduleType = moduleTypeOptions.find(
                    (option) => option.type === section.sectionType,
                  )

                  return (
                    <article
                      key={section.clientId}
                      className={[
                        styles.moduleCard,
                        draggedSectionClientId === section.clientId ? styles.moduleCardDragging : '',
                        dragOverSectionClientId === section.clientId &&
                        draggedSectionClientId !== section.clientId
                          ? styles.moduleCardDragOver
                          : '',
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      draggable={!isReadOnlyMode && !isBusy}
                      onDragStart={() => handleSectionDragStart(section.clientId)}
                      onDragOver={(event) => handleSectionDragOver(event, section.clientId)}
                      onDrop={() => handleSectionDrop(section.clientId)}
                      onDragEnd={handleSectionDragEnd}
                    >
                      <div className={styles.moduleHeader}>
                        <div className={styles.moduleHeaderMain}>
                          {!isReadOnlyMode && (
                            <button
                              type="button"
                              className={styles.dragHandle}
                              aria-label={t('pages.modules.actions.dragToReorder')}
                            >
                              <DragHandleIcon />
                            </button>
                          )}

                          <div className={styles.moduleHeadingBlock}>
                            <div className={styles.moduleEyebrowRow}>
                              <span className={styles.moduleEyebrow}>
                                {moduleType?.label ?? section.sectionType}
                              </span>
                              {!section.isEnabled && (
                                <span className={styles.moduleBadge}>
                                  {t('pages.modules.badges.hidden')}
                                </span>
                              )}
                            </div>
                            <h3 className={styles.moduleTitle}>
                              {section.sectionName || moduleType?.label}
                            </h3>
                          </div>
                        </div>

                        <div className={styles.moduleHeaderActions}>
                          <label className={styles.moduleToggle}>
                            <span>{t('pages.modules.fields.enabled')}</span>
                            <input
                              type="checkbox"
                              checked={section.isEnabled}
                              onChange={(event) =>
                                updateSection(section.clientId, (current) => ({
                                  ...current,
                                  isEnabled: event.target.checked,
                                }))
                              }
                            />
                          </label>

                          {!isReadOnlyMode && (
                            <>
                              <button
                                type="button"
                                className={styles.iconButton}
                                disabled={sectionIndex === 0}
                                aria-label={t('pages.modules.actions.moveUp')}
                                onClick={() => handleMoveSection(section.clientId, -1)}
                              >
                                <ArrowUpIcon />
                              </button>
                              <button
                                type="button"
                                className={styles.iconButton}
                                disabled={sectionIndex === form.sections.length - 1}
                                aria-label={t('pages.modules.actions.moveDown')}
                                onClick={() => handleMoveSection(section.clientId, 1)}
                              >
                                <ArrowDownIcon />
                              </button>
                            </>
                          )}

                          <button
                            type="button"
                            className={styles.iconButton}
                            aria-label={
                              section.isCollapsed
                                ? t('pages.modules.actions.expand')
                                : t('pages.modules.actions.collapse')
                            }
                            onClick={() =>
                              updateSection(section.clientId, (current) => ({
                                ...current,
                                isCollapsed: !current.isCollapsed,
                              }))
                            }
                          >
                            {section.isCollapsed ? <ChevronDownIcon /> : <ChevronUpIcon />}
                          </button>

                          {!isReadOnlyMode && (
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label={t('pages.modules.actions.removeSection')}
                              onClick={() => handleRemoveSection(section.clientId)}
                            >
                              <DeleteIcon size={16} />
                            </button>
                          )}
                        </div>
                      </div>

                      {!section.isCollapsed && (
                        <div className={styles.moduleBody}>
                          {renderSectionBody(section)}
                        </div>
                      )}
                    </article>
                  )
                })}
              </div>
            )}

            {!isReadOnlyMode && (
              <div className={styles.modulePickerWrap}>
                <button
                  type="button"
                  className={styles.primaryButton}
                  onClick={() => setModulePickerOpen((current) => !current)}
                >
                  <AddIcon size={16} />
                  {t('pages.modules.actions.addModule')}
                </button>

                {modulePickerOpen && (
                  <div className={styles.modulePicker}>
                    <p className={styles.modulePickerLabel}>
                      {t('pages.modules.actions.selectComponent')}
                    </p>
                    <div className={styles.modulePickerGrid}>
                      {moduleTypeOptions.map((option) => (
                        <button
                          key={option.type}
                          type="button"
                          className={styles.modulePickerItem}
                          onClick={() => handleAddSection(option.type)}
                        >
                          <span className={styles.modulePickerIcon}>{option.icon}</span>
                          <span className={styles.modulePickerText}>
                            <strong>{option.label}</strong>
                            <span>{option.description}</span>
                          </span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </section>
        </fieldset>

        {!isReadOnlyMode && (
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

function SegmentedButton({
  active,
  label,
  onClick,
}: {
  active: boolean
  label: string
  onClick: () => void
}) {
  return (
    <button
      type="button"
      className={[styles.segmentedButton, active ? styles.segmentedButtonActive : '']
        .filter(Boolean)
        .join(' ')}
      onClick={onClick}
    >
      {label}
    </button>
  )
}

function moveItemByClientId<T extends { clientId: string }>(
  items: T[],
  clientId: string,
  direction: -1 | 1,
) {
  const fromIndex = items.findIndex((item) => item.clientId === clientId)
  const toIndex = fromIndex + direction
  if (fromIndex < 0 || toIndex < 0 || toIndex >= items.length) {
    return items
  }

  const next = items.slice()
  const [moved] = next.splice(fromIndex, 1)
  next.splice(toIndex, 0, moved)
  return next
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

function DragHandleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <circle cx="5" cy="4" r="1.1" fill="currentColor" />
      <circle cx="11" cy="4" r="1.1" fill="currentColor" />
      <circle cx="5" cy="8" r="1.1" fill="currentColor" />
      <circle cx="11" cy="8" r="1.1" fill="currentColor" />
      <circle cx="5" cy="12" r="1.1" fill="currentColor" />
      <circle cx="11" cy="12" r="1.1" fill="currentColor" />
    </svg>
  )
}

function ArrowUpIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M8 12V4M8 4 4.8 7.2M8 4l3.2 3.2"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function ArrowDownIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M8 4v8M8 12l-3.2-3.2M8 12l3.2-3.2"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function ChevronUpIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="m4 10 4-4 4 4"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function ChevronDownIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="m4 6 4 4 4-4"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function TypographyModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M3 4h10M8 4v8M5.2 12h5.6"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function GalleryModuleIcon() {
  return <GalleryIcon size={16} />
}

function DocumentModuleIcon() {
  return <PagesIcon size={16} />
}

function QuoteModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M6.2 5.2H4.7A1.7 1.7 0 0 0 3 6.9v1.7a1.7 1.7 0 0 0 1.7 1.7h1.5A1.7 1.7 0 0 0 7.9 8.6V6.9A1.7 1.7 0 0 0 6.2 5.2ZM11.3 5.2H9.8A1.7 1.7 0 0 0 8.1 6.9v1.7a1.7 1.7 0 0 0 1.7 1.7h1.5A1.7 1.7 0 0 0 13 8.6V6.9a1.7 1.7 0 0 0-1.7-1.7Z"
        stroke="currentColor"
        strokeWidth="1.2"
      />
    </svg>
  )
}

function CtaModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <rect x="2.5" y="4" width="11" height="8" rx="1.7" stroke="currentColor" strokeWidth="1.2" />
      <path
        d="M5 8h6M9 6.2 11 8 9 9.8"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function HeaderModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M4 3.5v9M12 3.5v9M4 8h8"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
      />
    </svg>
  )
}
